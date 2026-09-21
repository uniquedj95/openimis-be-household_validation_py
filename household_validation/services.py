import json
from dataclasses import dataclass, replace
from datetime import date
from uuid import UUID, uuid4

from django.core.exceptions import ValidationError
from django.core.serializers.json import DjangoJSONEncoder
from django.db import transaction
from django.db.models import Q
from django.utils import timezone

from individual.models import Group, GroupIndividual
from individual.services import GroupIndividualService, IndividualService
from location.models import Hotspot, MicroCatchment
from project_social_protection.models import Project

from household_validation.excel import (
    HAS_BUSINESS_COLUMN,
    LOCATION_COLUMN_TYPES,
    is_primary_worker_rejection,
)
from household_validation.models import (
    HouseholdValidationBatch,
    HouseholdValidationBatchRow,
)
from household_validation.project_lookup import (
    ACTIVE_PROJECT_STATUSES,
    project_option_from_project,
)
from household_validation.selection import (
    EligibleHousehold,
    EligibleMember,
    is_truthy,
    select_households,
)
from household_validation.upload import (
    build_validation_error_report_csv,
    build_validation_json_ext,
    member_structural_errors,
    parse_validation_workbook,
)
from household_validation.wealth import get_household_pmt_score, get_household_wealth_quintile
from household_validation.verification import (
    VERIFIED, NOT_VERIFIED, REJECTED, BUSINESS_REJECTION_CODE,
    participant_status, household_status,
)


def _local_date(value):
    """Local date of ``value``, tolerating naive datetimes.

    openIMIS runs with ``USE_TZ = False``, so ``timezone.now()`` is naive and
    ``timezone.localdate()``/``timezone.localtime()`` raise ValueError. Mirrors
    the guard used in ``core.services.userServices``.
    """
    return timezone.localtime(value).date() if timezone.is_aware(value) else value.date()


def _json_safe(value):
    """Return a JSON-native copy suitable for storage in a JSONField."""
    return json.loads(json.dumps(value, cls=DjangoJSONEncoder))


class HouseholdValidationUploadService:
    def __init__(self, user=None):
        from household_validation.apps import HouseholdValidationConfig

        self.user = user
        self.business_columns_enabled = HouseholdValidationConfig.business_columns_enabled
        self._group_cache = {}
        self._upload_attempt_id = None
        self._member_details_changed_group_ids = set()

    def upload(self, file_or_bytes, dry_run=False, source_file_name=None):
        self._group_cache = {}
        self._member_details_changed_group_ids = set()
        self._upload_attempt_id = None
        parsed = parse_validation_workbook(file_or_bytes)
        totals = {
            "rows_read": parsed.rows_read,
            "households_verified": 0,
            "participants_verified": 0,
            "participants_not_verified": 0,
            "participants_rejected": 0,
            "households_not_verified": 0,
            "households_rejected": 0,
            "participant_updates": 0,
            "households_with_multiple_primary_workers": 0,
            "errors": len(parsed.errors),
            "error_messages": list(parsed.errors),
        }
        participant_update_group_keys = (
            self._accepted_validation_group_keys(parsed.rows)
            - set(getattr(parsed, "invalid_group_keys", ()))
        )
        primary_worker_rejections = self._primary_worker_rejections(
            parsed.rows,
            eligible_group_keys=participant_update_group_keys,
        )
        totals["households_with_multiple_primary_workers"] = len(
            primary_worker_rejections
        )
        totals["participants_rejected"] = len(
            {
                str(uploaded_row.values["member_uuid"]).strip()
                for uploaded_row in parsed.rows
                if self._uploaded_group_key(uploaded_row)
                in primary_worker_rejections
                and uploaded_row.primary_worker is True
            }
        )
        decision_rows = {}
        for row in parsed.rows:
            group_key = self._uploaded_group_key(row)
            if group_key in participant_update_group_keys:
                decision_rows.setdefault(group_key, []).append(
                    (row.primary_worker, row.values.get(HAS_BUSINESS_COLUMN))
                )
        decisions = {
            key: household_status(rows, business_columns_enabled=self.business_columns_enabled)
            for key, rows in decision_rows.items()
        }
        totals["households_rejected"] = sum(
            status == REJECTED for status in decisions.values()
        )
        verification_statuses = self._primary_worker_verification_statuses(
            parsed.rows,
            eligible_group_keys=participant_update_group_keys,
            rejected_group_keys=set(primary_worker_rejections),
        )
        uploaded_rows = [
            replace(
                uploaded_row,
                verified=verification_statuses.get(
                    self._uploaded_group_key(uploaded_row)
                ),
                household_status=decisions.get(self._uploaded_group_key(uploaded_row)),
                participant_status=(
                    participant_status(
                        uploaded_row.primary_worker,
                        uploaded_row.values.get(HAS_BUSINESS_COLUMN),
                        business_columns_enabled=self.business_columns_enabled,
                    )
                    if self._uploaded_group_key(uploaded_row) in participant_update_group_keys
                    else None
                ),
            )
            for uploaded_row in parsed.rows
        ]
        uploaded_rows_by_group = {}
        for uploaded_row in uploaded_rows:
            uploaded_rows_by_group.setdefault(
                self._uploaded_group_key(uploaded_row),
                [],
            ).append(uploaded_row)

        totals["participants_rejected"] += len({
            str(row.values["member_uuid"]).strip() for row in uploaded_rows
            if row.participant_status == REJECTED
            and self._uploaded_group_key(row) not in primary_worker_rejections
        })
        if dry_run:
            totals["households_verified"] = sum(status == VERIFIED for status in decisions.values())
            totals["households_not_verified"] = sum(status == NOT_VERIFIED for status in decisions.values())
            for status, key in ((VERIFIED, "participants_verified"), (NOT_VERIFIED, "participants_not_verified")):
                totals[key] = len({
                    str(row.values["member_uuid"]).strip() for row in uploaded_rows
                    if row.participant_status == status
                    and self._uploaded_group_key(row) not in primary_worker_rejections
                })
            return totals

        self._upload_attempt_id = uuid4()
        batch = self._get_or_create_batch(
            parsed,
            source_file_name=source_file_name,
        )
        totals["batch_id"] = batch.id
        totals["upload_attempt_id"] = self._upload_attempt_id
        uploaded_at = timezone.now()
        upload_date = _local_date(uploaded_at)
        verified_group_ids = set()
        verified_participant_ids = set()
        not_verified_group_ids = set()
        not_verified_participant_ids = set()
        primary_worker_changed_group_ids = set()
        national_id_changed_group_ids = set()
        successful_rows_by_group = {}

        with transaction.atomic():
            for group_key, verified in verification_statuses.items():
                if verified is True and self._replace_primary_worker(
                    uploaded_rows_by_group[group_key]
                ):
                    primary_worker_changed_group_ids.add(group_key)

            for uploaded_row in uploaded_rows:
                group_key = self._uploaded_group_key(uploaded_row)
                primary_worker_rejection = primary_worker_rejections.get(group_key)
                if primary_worker_rejection:
                    self._save_primary_worker_rejection(
                        uploaded_row,
                        batch=batch,
                    )
                    continue
                row_errors, national_id_updated = self._apply_row(
                    uploaded_row,
                    batch=batch,
                    upload_date=upload_date,
                    uploaded_at=uploaded_at,
                    allow_participant_update=group_key in participant_update_group_keys,
                )
                if row_errors:
                    totals["errors"] += len(row_errors)
                    totals["error_messages"].extend(row_errors)
                    continue
                successful_rows_by_group.setdefault(group_key, []).append(
                    uploaded_row
                )
                if uploaded_row.household_status == VERIFIED:
                    verified_group_ids.add(group_key)
                elif uploaded_row.household_status == NOT_VERIFIED:
                    not_verified_group_ids.add(group_key)
                if uploaded_row.participant_status == VERIFIED:
                    verified_participant_ids.add(str(uploaded_row.values["member_uuid"]).strip())
                elif uploaded_row.participant_status == NOT_VERIFIED:
                    not_verified_participant_ids.add(
                        str(uploaded_row.values["member_uuid"]).strip()
                    )
                if national_id_updated:
                    totals["participant_updates"] += 1
                    national_id_changed_group_ids.add(group_key)

            for group_key, group_rows in successful_rows_by_group.items():
                verified = verification_statuses.get(group_key)
                if verified is None:
                    continue
                representative_row = next(
                    (
                        row
                        for row in group_rows
                        if row.primary_worker is True
                    ),
                    group_rows[0],
                )
                self._apply_group_validation_if_changed(
                    group=self._group(group_key),
                    uploaded_row=representative_row,
                    project=self._project(representative_row.project_id),
                    upload_date=upload_date,
                    uploaded_at=uploaded_at,
                    force=(
                        group_key in primary_worker_changed_group_ids
                        or group_key in national_id_changed_group_ids
                        or group_key in self._member_details_changed_group_ids
                    ),
                )

            totals["households_verified"] = len(verified_group_ids)
            totals["participants_verified"] = len(verified_participant_ids)
            totals["households_not_verified"] = len(not_verified_group_ids)
            totals["participants_not_verified"] = len(
                not_verified_participant_ids
            )
            batch.uploaded_at = uploaded_at
            batch.status = self._batch_status(totals)
            batch.error_summary = "\n".join(totals["error_messages"]) or None
            batch.save(user=self.user)
        return totals

    def _accepted_validation_group_keys(self, uploaded_rows):
        group_is_valid = {}
        for uploaded_row in uploaded_rows:
            errors, _, _, _ = self._resolve_row(uploaded_row)
            group_key = self._uploaded_group_key(uploaded_row)
            group_is_valid[group_key] = (
                group_is_valid.get(group_key, True) and not errors
            )
        return {
            group_key
            for group_key, is_valid in group_is_valid.items()
            if is_valid
        }

    def _primary_worker_verification_statuses(
        self,
        uploaded_rows,
        eligible_group_keys,
        rejected_group_keys=None,
    ):
        """Derive household verification from its projected primary workers."""
        rejected_group_keys = rejected_group_keys or set()
        rows_by_group = {}
        for uploaded_row in uploaded_rows:
            rows_by_group.setdefault(
                self._uploaded_group_key(uploaded_row),
                [],
            ).append(uploaded_row)

        statuses = {}
        for group_key, group_rows in rows_by_group.items():
            if (
                group_key not in eligible_group_keys
                or group_key in rejected_group_keys
            ):
                continue
            projected = self._projected_primary_workers(group_rows)
            if projected is not None:
                statuses[group_key] = household_status([
                    (row.primary_worker, row.values.get(HAS_BUSINESS_COLUMN))
                    for row in group_rows
                ], business_columns_enabled=self.business_columns_enabled) == VERIFIED
        return statuses

    def _primary_worker_rejections(
        self,
        uploaded_rows,
        eligible_group_keys=None,
    ):
        rows_by_group = {}
        for uploaded_row in uploaded_rows:
            rows_by_group.setdefault(
                self._uploaded_group_key(uploaded_row),
                [],
            ).append(uploaded_row)

        rejections = {}
        for group_key, group_rows in rows_by_group.items():
            if (
                eligible_group_keys is not None
                and group_key not in eligible_group_keys
            ):
                continue
            projected = self._projected_primary_workers(group_rows)
            if projected is None:
                continue

            primary_worker_count = sum(projected.values())
            if primary_worker_count <= 1:
                continue

            rejections[group_key] = {
                "row_count": len(group_rows),
                "row_numbers": {
                    uploaded_row.row_number for uploaded_row in group_rows
                },
            }
        return rejections

    def _projected_primary_workers(self, group_rows):
        group = self._group(group_rows[0].values["group_uuid"])
        if group is None:
            return None
        # The workbook is the complete validation decision for the selected
        # participants. Stored Primary Worker flags must not turn blank cells
        # into implicit selections.
        projected = {}
        for uploaded_row in group_rows:
            group_individual = self._group_individual(
                uploaded_row.values["member_uuid"],
                group=group,
            )
            if group_individual is None:
                continue
            if self._structural_errors(
                uploaded_row,
                group,
                group_individual=group_individual,
            ):
                continue
            if (
                uploaded_row.project_id
                and self._project(uploaded_row.project_id) is None
            ):
                continue
            projected[str(group_individual.individual_id)] = (
                uploaded_row.primary_worker is True
            )
        return projected

    def _replace_primary_worker(self, group_rows):
        """Replace a household's assignment from its single workbook YES."""
        selected_rows = [
            uploaded_row
            for uploaded_row in group_rows
            if uploaded_row.primary_worker is True
        ]
        if len(selected_rows) != 1:
            raise ValidationError(
                "A Primary Worker replacement requires exactly one selected participant"
            )

        group = self._group(selected_rows[0].values["group_uuid"])
        selected_member_id = str(selected_rows[0].values["member_uuid"])
        active_members = list(
            GroupIndividual.objects.select_for_update().filter(
                group=group,
                is_deleted=False,
            )
        )
        selected_members = [
            member
            for member in active_members
            if str(member.individual_id) == selected_member_id
        ]
        if len(selected_members) != 1:
            raise ValidationError(
                "The selected Primary Worker is not an active household member"
            )

        # Clear every previous assignment first, including members that were
        # not present in the workbook, then assign the selected participant.
        changed = False
        for member in active_members:
            if (
                str(member.individual_id) != selected_member_id
                and self._primary_worker_changed(member, False)
            ):
                self._apply_primary_worker(member, False)
                changed = True

        selected_member = selected_members[0]
        if self._primary_worker_changed(selected_member, True):
            self._apply_primary_worker(selected_member, True)
            changed = True
        return changed

    def _save_primary_worker_rejection(self, uploaded_row, batch):
        uploaded_row = replace(
            uploaded_row,
            verified=None,
            participant_status=REJECTED if uploaded_row.primary_worker is True else None,
            household_status=REJECTED,
        )
        group = self._group(uploaded_row.values["group_uuid"])
        group_individual = self._group_individual(
            uploaded_row.values["member_uuid"],
            group=group,
        )
        self._save_batch_row(
            batch=batch,
            uploaded_row=uploaded_row,
            group=group,
            group_individual=group_individual,
            project=self._project(uploaded_row.project_id),
            status=HouseholdValidationBatchRow.Status.REJECTED,
            error_message="household has more than one primary worker",
            error_code="MULTIPLE_PRIMARY_WORKERS",
        )

    @staticmethod
    def _uploaded_group_key(uploaded_row):
        return str(uploaded_row.values.get("group_uuid") or "").strip()

    def _get_or_create_batch(self, parsed, source_file_name=None):
        batch_id = self._batch_id(parsed)
        if batch_id:
            batch = HouseholdValidationBatch.objects.filter(id=batch_id).first()
            if batch:
                return batch
        batch = HouseholdValidationBatch(
            source_file_name=source_file_name,
            status=HouseholdValidationBatch.Status.PENDING,
        )
        if batch_id:
            batch.id = batch_id
        batch.save(user=self.user)
        return batch

    def _batch_id(self, parsed):
        batch_ids = {
            row.values.get("batch_id")
            for row in parsed.rows
            if row.values.get("batch_id")
        }
        if len(batch_ids) != 1:
            return None
        batch_id = next(iter(batch_ids))
        try:
            return UUID(str(batch_id))
        except ValueError:
            return None

    def _resolve_row(self, uploaded_row):
        errors = []
        group = self._group(uploaded_row.values["group_uuid"])
        group_individual = self._group_individual(
            uploaded_row.values["member_uuid"],
            group=group,
        )
        project = self._project(uploaded_row.project_id)

        if group is None:
            errors.append(f"Row {uploaded_row.row_number}: group was not found")
        else:
            errors.extend(
                self._structural_errors(
                    uploaded_row,
                    group,
                    group_individual=group_individual,
                )
            )
        if group_individual is None:
            errors.append(f"Row {uploaded_row.row_number}: member was not found in group")
        if uploaded_row.project_id and project is None:
            errors.append(f"Row {uploaded_row.row_number}: project was not found")

        return errors, group, group_individual, project

    def _apply_row(
        self,
        uploaded_row,
        batch,
        upload_date,
        uploaded_at,
        allow_participant_update=False,
    ):
        errors, group, group_individual, project = self._resolve_row(uploaded_row)

        if errors:
            self._save_batch_row(
                batch=batch,
                uploaded_row=uploaded_row,
                group=group,
                group_individual=group_individual,
                project=project,
                status=HouseholdValidationBatchRow.Status.ERROR,
                error_message="\n".join(errors),
            )
            return errors, False

        national_id_updated = (
            allow_participant_update
            and "national_id" in uploaded_row.values
            and self._national_id_changed(
                group_individual,
                uploaded_row.values.get("national_id"),
            )
        )
        if national_id_updated:
            self._apply_national_id(
                group_individual,
                uploaded_row.values.get("national_id"),
            )
        details_updated = False
        if allow_participant_update:
            details_updated = self._apply_member_details(group_individual, uploaded_row)
            if details_updated:
                self._member_details_changed_group_ids.add(self._uploaded_group_key(uploaded_row))
        self._save_batch_row(
            batch=batch,
            uploaded_row=uploaded_row,
            group=group,
            group_individual=group_individual,
            project=project,
            status=(
                HouseholdValidationBatchRow.Status.REJECTED
                if uploaded_row.household_status == REJECTED
                else
                HouseholdValidationBatchRow.Status.APPLIED
                if (
                    uploaded_row.verified is not None
                    or national_id_updated
                    or details_updated
                )
                else HouseholdValidationBatchRow.Status.SKIPPED
            ),
            error_code=(BUSINESS_REJECTION_CODE if uploaded_row.household_status == REJECTED else None),
            error_message=(
                "household has a business-owning member without Primary Worker = Yes"
                if uploaded_row.household_status == REJECTED else None
            ),
        )
        # The public participant_updates result is labelled National IDs
        # Updated. Primary Worker changes are deliberately excluded.
        return [], national_id_updated

    @staticmethod
    def _primary_worker_changed(group_individual, primary_worker):
        current_value = is_truthy(
            (group_individual.json_ext or {}).get("primary_worker")
        )
        return current_value != primary_worker

    @classmethod
    def _national_id_changed(cls, group_individual, national_id):
        individual = getattr(group_individual, "individual", None)
        if individual is None:
            return False
        current_value = (individual.json_ext or {}).get("national_id")
        return cls._normalize_national_id(current_value) != cls._normalize_national_id(
            national_id
        )

    @staticmethod
    def _normalize_national_id(value):
        if value is None:
            return None
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        value = str(value).strip()
        return value or None

    def _apply_group_validation_if_changed(
        self,
        group,
        uploaded_row,
        project,
        upload_date,
        uploaded_at,
        force=False,
    ):
        json_ext = dict(group.json_ext or {})
        candidate = build_validation_json_ext(
            uploaded_row=uploaded_row,
            project=project,
            upload_date=upload_date,
            uploaded_at=uploaded_at,
            user_id=getattr(self.user, "id", None),
        )
        meaningful_keys = (
            "validation_status",
            "validation_project_id",
            "validation_project_name",
            "validation_project_selection_type",
            "validation_notes",
        )
        changed = (
            force
            or not json_ext.get("last_verified_date")
            or any(
                json_ext.get(key) != candidate.get(key)
                for key in meaningful_keys
            )
        )
        if not changed:
            return False
        json_ext.update(candidate)
        group.json_ext = json_ext
        group.save(user=self.user)
        return True

    def _apply_primary_worker(self, group_individual, primary_worker):
        json_ext = dict(group_individual.json_ext or {})
        json_ext["primary_worker"] = primary_worker
        GroupIndividualService(self.user).update(
            {
                "id": group_individual.id,
                "group_id": group_individual.group_id,
                "json_ext": json_ext,
            }
        )

    def _apply_national_id(self, group_individual, national_id):
        individual = group_individual.individual
        json_ext = dict(individual.json_ext or {})
        json_ext["national_id"] = self._normalize_national_id(national_id)
        IndividualService(self.user).update(
            {
                "id": individual.id,
                "json_ext": json_ext,
            }
        )
        individual.json_ext = json_ext

    def _apply_member_details(self, group_individual, uploaded_row):
        updates = dict(uploaded_row.business_updates)
        if "validation_notes" in uploaded_row.values:
            updates["validation_notes"] = uploaded_row.notes
        if uploaded_row.participant_status:
            updates["validation_status"] = uploaded_row.participant_status
        if not updates:
            return False
        individual = group_individual.individual
        json_ext = dict(individual.json_ext or {})
        if not any(json_ext.get(key) != value for key, value in updates.items()):
            return False
        json_ext.update(updates)
        IndividualService(self.user).update({"id": individual.id, "json_ext": json_ext})
        individual.json_ext = json_ext
        return True

    def _save_batch_row(
        self,
        batch,
        uploaded_row,
        group=None,
        group_individual=None,
        project=None,
        status=HouseholdValidationBatchRow.Status.PENDING,
        error_message=None,
        error_code=None,
    ):
        batch_row = HouseholdValidationBatchRow(
            batch=batch,
            upload_attempt_id=self._upload_attempt_id,
            group=group,
            group_individual=group_individual,
            individual=getattr(group_individual, "individual", None),
            project=project,
            row_number=uploaded_row.row_number,
            verified=(
                uploaded_row.participant_status == VERIFIED
                if uploaded_row.participant_status else uploaded_row.verified
            ),
            validation_date=uploaded_row.validation_date,
            status=status,
            error_message=error_message,
            raw_row=_json_safe(uploaded_row.values),
            json_ext={
                "primary_worker": uploaded_row.primary_worker,
                "project_name": uploaded_row.project_name,
                "validation_notes": uploaded_row.notes,
                "error_code": error_code,
                "participant_status": uploaded_row.participant_status,
                "household_status": uploaded_row.household_status,
            },
        )
        batch_row.save(user=self.user)

    def _project(self, project_id):
        if not project_id:
            return None
        try:
            return Project.objects.filter(id=project_id).first()
        except (ValueError, ValidationError):
            return None

    def _group(self, group_id):
        cache_key = str(group_id)
        if cache_key in self._group_cache:
            return self._group_cache[cache_key]
        try:
            group = (
                Group.objects.filter(id=group_id)
                .select_related("location")
                .prefetch_related("groupindividuals__individual")
                .first()
            )
        except (ValueError, ValidationError):
            group = None
        self._group_cache[cache_key] = group
        return group

    def _group_individual(self, individual_id, group):
        prefetched_members = getattr(
            group,
            "_prefetched_objects_cache",
            {},
        ).get("groupindividuals")
        if prefetched_members is not None:
            for group_individual in prefetched_members:
                if (
                    str(group_individual.individual_id) == str(individual_id)
                    and not group_individual.is_deleted
                ):
                    return group_individual
            return None
        try:
            return (
                GroupIndividual.objects.filter(
                    individual_id=individual_id,
                    group=group,
                    is_deleted=False,
                )
                .select_related("individual")
                .first()
            )
        except (ValueError, ValidationError):
            return None

    def _structural_errors(self, uploaded_row, group, group_individual=None):
        errors = []
        row_number = uploaded_row.row_number

        if uploaded_row.values.get("row_type") not in ("MAIN", "RESERVE"):
            errors.append(f"Row {row_number}: row_type is invalid")
        location = getattr(group, "location", None)
        for column, location_type in LOCATION_COLUMN_TYPES.items():
            uploaded_value = self._normalize(uploaded_row.values.get(column))
            database_value = self._normalize(self._location_name(location, location_type))
            if uploaded_value and database_value and uploaded_value != database_value:
                errors.append(f"Row {row_number}: {column} does not match the household")
        if group_individual is not None:
            errors.extend(
                member_structural_errors(
                    uploaded_row,
                    group_individual=group_individual,
                    group=group,
                )
            )
        return errors

    def _location_name(self, location, location_type):
        current = location
        while current is not None:
            if getattr(current, "type", None) == location_type:
                return getattr(current, "name", None) or getattr(current, "code", None)
            current = getattr(current, "parent", None)
        return None

    def _normalize(self, value):
        if value is None:
            return None
        if isinstance(value, date):
            value = value.isoformat()
        value = str(value).strip().casefold()
        return value or None

    def _batch_status(self, totals):
        if totals["errors"] and totals["errors"] >= totals["rows_read"]:
            return HouseholdValidationBatch.Status.FAILED
        if totals["errors"]:
            return HouseholdValidationBatch.Status.PARTIAL_SUCCESS
        return HouseholdValidationBatch.Status.PROCESSED


def build_validation_error_report(batch):
    rows = [
        row
        for row in batch.rows.filter(
            status=HouseholdValidationBatchRow.Status.ERROR,
            is_deleted=False,
        ).order_by("row_number")
        if not is_primary_worker_rejection(row)
    ]
    return build_validation_error_report_csv(batch.id, rows)


class HouseholdValidationProjectLookupService:
    def list_projects(
        self,
        location_id=None,
        location_code=None,
        hotspot_id=None,
        hotspot_code=None,
        catchment_id=None,
    ):
        queryset = Project.objects.select_related("location").filter(
            status__in=ACTIVE_PROJECT_STATUSES,
        )
        queryset = self._apply_location_filter(
            queryset,
            location_id=location_id,
            location_code=location_code,
        )
        # Hotspot and public works catchment models do not exist yet. Keep the
        # nullable filter signature stable and add relation filters when they do.
        return [
            project_option_from_project(project)
            for project in queryset.order_by("name", "id")
        ]

    def _apply_location_filter(self, queryset, location_id=None, location_code=None):
        location_filter = Q()
        if location_id:
            location_filter |= (
                Q(location_id=location_id)
                | Q(location__parent_id=location_id)
                | Q(location__parent__parent_id=location_id)
                | Q(location__children__id=location_id)
                | Q(location__children__children__id=location_id)
            )
        if location_code:
            location_filter |= (
                Q(location__code=location_code)
                | Q(location__parent__code=location_code)
                | Q(location__parent__parent__code=location_code)
                | Q(location__children__code=location_code)
                | Q(location__children__children__code=location_code)
            )
        if not location_filter:
            return queryset
        return queryset.filter(location_filter).distinct()


@dataclass(frozen=True)
class HouseholdValidationPreviewRow:
    row_type: str
    category: str
    group_uuid: str
    group_code: str | None
    head_name: str | None
    individual_uuid: str
    individual_first_name: str | None
    individual_last_name: str | None
    individual_dob: date | None
    individual_age: int | None
    individual_gender: str | None
    fit_for_work: bool
    current_recipient_type: str | None
    region: str | None
    district: str | None
    municipality: str | None
    village: str | None
    wealth_quintile: str | int | None
    last_verified_date: date | None
    validation_status: str | None
    prospective_projects: list[str]


class EligibleHouseholdSelectionService:
    def __init__(self, user=None):
        self.user = user
        self._project_name_cache = {}

    def select(
        self,
        region_id=None,
        region_code=None,
        district_id=None,
        district_code=None,
        ta_id=None,
        ta_code=None,
        ta_codes=None,
        gvh_codes=None,
        village_id=None,
        village_code=None,
        village_codes=None,
        hotspot_id=None,
        hotspot_code=None,
        catchment_id=None,
        catchment_code=None,
        exclude_verified_after=None,
        target_count=None,
    ):
        candidates = self.candidates(
            region_id=region_id,
            region_code=region_code,
            district_id=district_id,
            district_code=district_code,
            ta_id=ta_id,
            ta_code=ta_code,
            ta_codes=ta_codes,
            gvh_codes=gvh_codes,
            village_id=village_id,
            village_code=village_code,
            village_codes=village_codes,
            hotspot_id=hotspot_id,
            hotspot_code=hotspot_code,
            catchment_id=catchment_id,
            catchment_code=catchment_code,
        )
        selection_result, _ = select_households(
            candidates,
            target_count=target_count,
            exclude_verified_after=exclude_verified_after,
            allocate_by_village=bool(
                catchment_id
                or catchment_code
                or hotspot_id
                or hotspot_code
            ),
        )
        return selection_result

    def candidates(
        self,
        region_id=None,
        region_code=None,
        district_id=None,
        district_code=None,
        ta_id=None,
        ta_code=None,
        ta_codes=None,
        gvh_codes=None,
        village_id=None,
        village_code=None,
        village_codes=None,
        hotspot_id=None,
        hotspot_code=None,
        catchment_id=None,
        catchment_code=None,
    ):
        queryset = self._base_queryset()
        queryset = self._apply_location_filters(
            queryset,
            region_id=region_id,
            region_code=region_code,
            district_id=district_id,
            district_code=district_code,
            ta_id=ta_id,
            ta_code=ta_code,
            ta_codes=ta_codes,
            gvh_codes=gvh_codes,
            village_id=village_id,
            village_code=village_code,
            village_codes=village_codes,
            hotspot_id=hotspot_id,
            hotspot_code=hotspot_code,
            catchment_id=catchment_id,
            catchment_code=catchment_code,
        )
        return [
            household
            for household in (self._build_household(group) for group in queryset)
            if household is not None
        ]

    def generate(self, **filters):
        """
        The ``generateHouseholdValidationList`` mutation exposes both: the
        ``SelectionResult`` to export the workbook, and the summary counts inline
        on the same response, so callers get a single request/response.
        """
        base_queryset = self._base_queryset()
        catchment_id = filters.get("catchment_id")
        catchment_code = filters.get("catchment_code")

        # The headline totals always describe the complete selected micro-
        # catchment. Optional TA/GVH/hotspot/village filters narrow only the
        # validation selection and must not change those two totals.
        if catchment_id or catchment_code:
            catchment_queryset = self._apply_micro_catchment_scope(
                base_queryset,
                catchment_id=catchment_id,
                catchment_code=catchment_code,
            )
            queryset = self._apply_location_filters(
                catchment_queryset,
                region_id=filters.get("region_id"),
                region_code=filters.get("region_code"),
                district_id=filters.get("district_id"),
                district_code=filters.get("district_code"),
                ta_id=filters.get("ta_id"),
                ta_code=filters.get("ta_code"),
                ta_codes=filters.get("ta_codes"),
                gvh_codes=filters.get("gvh_codes"),
                village_id=filters.get("village_id"),
                village_code=filters.get("village_code"),
                village_codes=filters.get("village_codes"),
                hotspot_id=filters.get("hotspot_id"),
                hotspot_code=filters.get("hotspot_code"),
            )
            total_groups = list(catchment_queryset)
        else:
            queryset = self._apply_location_filters(
                base_queryset,
                region_id=filters.get("region_id"),
                region_code=filters.get("region_code"),
                district_id=filters.get("district_id"),
                district_code=filters.get("district_code"),
                ta_id=filters.get("ta_id"),
                ta_code=filters.get("ta_code"),
                ta_codes=filters.get("ta_codes"),
                gvh_codes=filters.get("gvh_codes"),
                village_id=filters.get("village_id"),
                village_code=filters.get("village_code"),
                village_codes=filters.get("village_codes"),
                hotspot_id=filters.get("hotspot_id"),
                hotspot_code=filters.get("hotspot_code"),
            )
            total_groups = list(queryset)

        groups = list(queryset)
        total_households = len(total_groups)
        total_individual_ids = {
            member.individual_id
            for group in total_groups
            for member in group.groupindividuals.all()
            if self._is_active_group_individual(member)
        }
        total_individuals = len(total_individual_ids)
        eligible_individuals = 0
        candidates = []
        for group in groups:
            household = self._build_household(group)
            if household is None:
                continue
            candidates.append(household)
            eligible_individuals += len(household.eligible_members)

        selection_result, selection_summary = select_households(
            candidates,
            target_count=filters.get("target_count"),
            exclude_verified_after=filters.get("exclude_verified_after"),
            allocate_by_village=bool(
                catchment_id
                or catchment_code
                or filters.get("hotspot_id")
                or filters.get("hotspot_code")
            ),
        )
        summary = {
            "total_households": total_households,
            "total_individuals": total_individuals,
            "eligible_households": len(candidates),
            "eligible_individuals": eligible_individuals,
            **selection_summary,
            "generated_at": timezone.now(),
        }
        return selection_result, summary

    def preview(self, **filters):
        selection_result = self.select(**filters)
        return [
            self._preview_row(selected_member)
            for selected_member in selection_result.member_rows
        ]

    def _base_queryset(self):
        queryset = (
            Group.objects.filter(is_deleted=False)
            .select_related("location")
            .prefetch_related("groupindividuals__individual")
        )
        if self.user is not None:
            queryset = Group.get_queryset(queryset, self.user)
        return queryset

    def _apply_location_filters(
        self,
        queryset,
        region_id=None,
        region_code=None,
        district_id=None,
        district_code=None,
        ta_id=None,
        ta_code=None,
        ta_codes=None,
        gvh_codes=None,
        village_id=None,
        village_code=None,
        village_codes=None,
        hotspot_id=None,
        hotspot_code=None,
        catchment_id=None,
        catchment_code=None,
    ):
        if catchment_id or catchment_code:
            queryset = self._apply_micro_catchment_scope(
                queryset,
                catchment_id=catchment_id,
                catchment_code=catchment_code,
            )

        # Accept both the legacy singular ``village_code`` and the plural
        # ``village_codes`` list; merge them into a single code set.
        codes = list(village_codes or [])
        if village_code:
            codes.append(village_code)

        if village_id or codes:
            village_filter = Q()
            if village_id:
                village_filter |= Q(location_id=village_id)
            if codes:
                village_filter |= Q(location__code__in=codes)
            return queryset.filter(village_filter)

        selected_gvh_codes = list(dict.fromkeys(gvh_codes or []))
        if selected_gvh_codes:
            return queryset.filter(
                Q(location__code__in=selected_gvh_codes)
                | Q(location__parent__code__in=selected_gvh_codes)
            )

        # A hotspot is a curated set of specific villages, finer-grained than a
        # plain TA/GVH pick, so it's checked before those but still yields to an
        # explicit village/GVH selection made on top of it in the UI.
        if hotspot_id or hotspot_code:
            hotspot = self._resolve_hotspot(hotspot_id, hotspot_code)
            if hotspot is None:
                return queryset.none()
            hotspot_village_codes = list(hotspot.villages.values_list("code", flat=True))
            if not hotspot_village_codes:
                return queryset.none()
            return queryset.filter(location__code__in=hotspot_village_codes)

        selected_ta_codes = list(dict.fromkeys(ta_codes or []))
        if ta_code:
            selected_ta_codes.append(ta_code)

        if ta_id or selected_ta_codes:
            ta_filter = Q()
            if ta_id:
                ta_filter |= (
                    Q(location_id=ta_id)
                    | Q(location__parent_id=ta_id)
                    | Q(location__parent__parent_id=ta_id)
                )
            if selected_ta_codes:
                ta_filter |= (
                    Q(location__code__in=selected_ta_codes)
                    | Q(location__parent__code__in=selected_ta_codes)
                    | Q(location__parent__parent__code__in=selected_ta_codes)
                )
            return queryset.filter(ta_filter)

        if district_id or district_code:
            district_filter = Q()
            if district_id:
                district_filter |= (
                    Q(location_id=district_id)
                    | Q(location__parent_id=district_id)
                    | Q(location__parent__parent_id=district_id)
                    | Q(location__parent__parent__parent_id=district_id)
                )
            if district_code:
                district_filter |= (
                    Q(location__code=district_code)
                    | Q(location__parent__code=district_code)
                    | Q(location__parent__parent__code=district_code)
                    | Q(location__parent__parent__parent__code=district_code)
                )
            return queryset.filter(district_filter)
        if region_id or region_code:
            region_filter = Q()
            if region_id:
                region_filter |= (
                    Q(location_id=region_id)
                    | Q(location__parent_id=region_id)
                    | Q(location__parent__parent_id=region_id)
                    | Q(location__parent__parent__parent_id=region_id)
                )
            if region_code:
                region_filter |= (
                    Q(location__code=region_code)
                    | Q(location__parent__code=region_code)
                    | Q(location__parent__parent__code=region_code)
                    | Q(location__parent__parent__parent__code=region_code)
                )
            return queryset.filter(region_filter)
        return queryset

    def _apply_micro_catchment_scope(
        self,
        queryset,
        catchment_id=None,
        catchment_code=None,
    ):
        micro_catchment = self._resolve_micro_catchment(catchment_id, catchment_code)
        if micro_catchment is None:
            return queryset.none()
        location_filter = self._micro_catchment_location_filter(micro_catchment)
        if location_filter is None:
            return queryset.none()
        return queryset.filter(location_filter)

    def _micro_catchment_location_filter(self, micro_catchment):
        """Return the most specific active location coverage for a catchment."""
        village_codes = list(
            Hotspot.objects.filter(
                micro_catchment=micro_catchment,
                validity_to__isnull=True,
                village_links__validity_to__isnull=True,
                village_links__location__validity_to__isnull=True,
            )
            .values_list("village_links__location__code", flat=True)
            .distinct()
        )
        if village_codes:
            return Q(location__code__in=village_codes)

        gvh_codes = list(
            micro_catchment.gvhs.filter(
                validity_to__isnull=True,
                location__validity_to__isnull=True,
            ).values_list("location__code", flat=True)
        )
        if gvh_codes:
            return (
                Q(location__code__in=gvh_codes)
                | Q(location__parent__code__in=gvh_codes)
            )

        ta_codes = list(
            micro_catchment.traditional_authorities.filter(
                validity_to__isnull=True,
                location__validity_to__isnull=True,
            ).values_list("location__code", flat=True)
        )
        if ta_codes:
            return (
                Q(location__code__in=ta_codes)
                | Q(location__parent__code__in=ta_codes)
                | Q(location__parent__parent__code__in=ta_codes)
            )
        return None

    def _resolve_hotspot(self, hotspot_id, hotspot_code):
        identity_filter = Q()
        if hotspot_id:
            identity_filter |= Q(uuid=hotspot_id)
        if hotspot_code:
            identity_filter |= Q(code=hotspot_code)
        if not identity_filter:
            return None
        return Hotspot.objects.filter(identity_filter, validity_to__isnull=True).first()

    def _resolve_micro_catchment(self, catchment_id, catchment_code):
        identity_filter = Q()
        if catchment_id:
            identity_filter |= Q(uuid=catchment_id)
        if catchment_code:
            identity_filter |= Q(code=catchment_code)
        if not identity_filter:
            return None
        return MicroCatchment.objects.filter(identity_filter, validity_to__isnull=True).first()

    def _build_household(self, group):
        groupindividuals = [
            group_individual
            for group_individual in group.groupindividuals.all()
            if self._is_active_group_individual(group_individual)
        ]
        eligible_members = []
        for group_individual in groupindividuals:
            member = self._build_member(group_individual)
            if member and member.fit_for_work:
                eligible_members.append(member)
        if not eligible_members:
            return None

        head = self._find_head(group, groupindividuals)
        wealth_quintile = get_household_wealth_quintile(
            group,
            group_individuals=groupindividuals,
        )
        pmt_score = get_household_pmt_score(
            group,
            group_individuals=groupindividuals,
        )

        group_json_ext = group.json_ext or {}
        location = getattr(group, "location", None)
        village = location if getattr(location, "type", None) == "V" else None

        return EligibleHousehold(
            id=group.id,
            code=group.code,
            wealth_quintile=wealth_quintile,
            pmt_score=pmt_score,
            last_verified_date=self._parse_date(group_json_ext.get("last_verified_date")),
            head=head,
            eligible_members=eligible_members,
            source=group,
            village_id=getattr(village, "id", None),
            village_code=getattr(village, "code", None),
            village_name=getattr(village, "name", None),
        )

    def _is_active_group_individual(self, group_individual):
        individual = getattr(group_individual, "individual", None)
        return (
            not getattr(group_individual, "is_deleted", False)
            and individual is not None
            and not getattr(individual, "is_deleted", False)
        )

    def _find_head(self, group, groupindividuals):
        head_id = (group.json_ext or {}).get("head_id")
        for group_individual in groupindividuals:
            if group_individual.role == GroupIndividual.Role.HEAD:
                return self._build_member(group_individual)
        if head_id:
            for group_individual in groupindividuals:
                if str(group_individual.individual_id) == str(head_id):
                    return self._build_member(group_individual)
        return None

    def _build_member(self, group_individual):
        individual = group_individual.individual
        if individual is None:
            return None
        json_ext = individual.json_ext or {}
        return EligibleMember(
            id=individual.id,
            gender=json_ext.get("gender"),
            dob=individual.dob,
            fit_for_work=self._is_fit_for_work(group_individual),
            role=group_individual.role,
            recipient_type=group_individual.recipient_type,
            source=group_individual,
        )

    def _is_fit_for_work(self, group_individual):
        individual = group_individual.individual
        if individual is None:
            return False
        return is_truthy((individual.json_ext or {}).get("fit_for_work"))

    def _parse_date(self, value):
        if isinstance(value, date):
            return value
        if not value:
            return None
        try:
            return date.fromisoformat(str(value)[:10])
        except ValueError:
            return None

    def _preview_row(self, selected_member):
        household = selected_member.household
        member = selected_member.member
        group = household.source
        group_individual = member.source
        individual = getattr(group_individual, "individual", None)
        location = getattr(group, "location", None)
        group_json_ext = getattr(group, "json_ext", None) or {}
        return HouseholdValidationPreviewRow(
            row_type=selected_member.row_type,
            category=selected_member.category,
            group_uuid=str(household.id),
            group_code=household.code,
            head_name=self._member_name(getattr(household.head, "source", None)),
            individual_uuid=str(member.id),
            individual_first_name=getattr(individual, "first_name", None),
            individual_last_name=getattr(individual, "last_name", None),
            individual_dob=getattr(individual, "dob", None),
            individual_age=member.age,
            individual_gender=member.gender,
            fit_for_work=member.fit_for_work,
            current_recipient_type=member.recipient_type,
            region=self._location_name(location, "R"),
            district=self._location_name(location, "D"),
            municipality=self._location_name(location, "W"),
            village=self._location_name(location, "V"),
            wealth_quintile=household.wealth_quintile,
            last_verified_date=household.last_verified_date,
            validation_status=group_json_ext.get("validation_status"),
            prospective_projects=self._project_names(location),
        )

    def _member_name(self, group_individual):
        individual = getattr(group_individual, "individual", None)
        if not individual:
            return None
        first_name = getattr(individual, "first_name", "") or ""
        last_name = getattr(individual, "last_name", "") or ""
        return f"{first_name} {last_name}".strip() or None

    def _location_name(self, location, location_type):
        current = location
        while current is not None:
            if getattr(current, "type", None) == location_type:
                return getattr(current, "name", None) or getattr(current, "code", None)
            current = getattr(current, "parent", None)
        return None

    def _project_names(self, location):
        location_id = getattr(location, "id", None)
        if not location_id:
            return []
        if location_id in self._project_name_cache:
            return self._project_name_cache[location_id]
        project_names = [
            project.name
            for project in HouseholdValidationProjectLookupService().list_projects(
                location_id=location_id,
            )
            if project.name
        ]
        self._project_name_cache[location_id] = project_names
        return project_names
