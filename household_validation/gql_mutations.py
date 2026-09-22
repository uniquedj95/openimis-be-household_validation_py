import base64

import graphene
from django.core.exceptions import PermissionDenied, ValidationError
from django.utils import timezone

from household_validation.apps import HouseholdValidationConfig
from household_validation.excel import ExcelValidationListExporter
from household_validation.gql_permissions import require_permissions
from household_validation.models import HouseholdValidationBatch
from household_validation.selection import (
    female_headed_percentage,
    reserve_percentage,
    youth_headed_percentage,
)
from household_validation.services import (
    EligibleHouseholdSelectionService,
    HouseholdValidationProjectLookupService,
    HouseholdValidationUploadService,
    _json_safe,
)


class HouseholdValidationVillageBreakdownGQLType(graphene.ObjectType):
    village_id = graphene.String()
    village_code = graphene.String()
    village_name = graphene.String()
    eligible_households = graphene.Int()
    exact_allocation = graphene.Float()
    allocated_households = graphene.Int()
    selected_households = graphene.Int()
    selected_individuals = graphene.Int()
    reserve_households = graphene.Int()


class HouseholdValidationGenerateResultGQLType(graphene.ObjectType):
    batch_id = graphene.UUID()
    file_name = graphene.String()
    file_base64 = graphene.String()
    total_households = graphene.Int()
    total_individuals = graphene.Int()
    eligible_households = graphene.Int()
    eligible_individuals = graphene.Int()
    selected_households = graphene.Int()
    selected_individuals = graphene.Int()
    selected_female_headed_households = graphene.Int()
    selected_youth_households = graphene.Int()
    selected_other_households = graphene.Int()
    reserve_households = graphene.Int()
    village_breakdown = graphene.List(HouseholdValidationVillageBreakdownGQLType)
    generated_at = graphene.DateTime()


class HouseholdValidationUploadResultGQLType(graphene.ObjectType):
    batch_id = graphene.UUID()
    upload_attempt_id = graphene.UUID()
    rows_read = graphene.Int()
    households_verified = graphene.Int()
    participants_verified = graphene.Int()
    participants_not_verified = graphene.Int()
    participants_rejected = graphene.Int()
    households_not_verified = graphene.Int()
    households_rejected = graphene.Int()
    participant_updates = graphene.Int()
    households_with_multiple_primary_workers = graphene.Int()
    errors = graphene.Int()
    error_messages = graphene.List(graphene.String)


class GenerateHouseholdValidationListMutation(graphene.Mutation):
    Output = HouseholdValidationGenerateResultGQLType

    class Arguments:
        region_id = graphene.Int(required=False)
        region_code = graphene.String(required=False)
        district_id = graphene.Int(required=False)
        district_code = graphene.String(required=False)
        ta_id = graphene.Int(required=False)
        ta_code = graphene.String(required=False)
        ta_codes = graphene.List(graphene.String, required=False)
        gvh_codes = graphene.List(graphene.String, required=False)
        village_id = graphene.Int(required=False)
        village_code = graphene.String(required=False)
        village_codes = graphene.List(graphene.String, required=False)
        hotspot_id = graphene.String(required=False)
        hotspot_code = graphene.String(required=False)
        catchment_id = graphene.String(required=False)
        catchment_code = graphene.String(required=False)
        exclude_verified_after = graphene.Date(required=False)
        target_count = graphene.Int(required=False)
        benefit_plan_code = graphene.String(required=False)

    @classmethod
    def mutate(cls, root, info, **data):
        cls._validate_user(
            info.context.user,
            HouseholdValidationConfig.gql_mutation_generate_household_validation_list_perms,
        )
        service = EligibleHouseholdSelectionService(info.context.user)
        selection_result, summary = service.generate(
            region_id=data.get("region_id"),
            region_code=data.get("region_code"),
            district_id=data.get("district_id"),
            district_code=data.get("district_code"),
            ta_id=data.get("ta_id"),
            ta_code=data.get("ta_code"),
            ta_codes=data.get("ta_codes"),
            gvh_codes=data.get("gvh_codes"),
            village_id=data.get("village_id"),
            village_code=data.get("village_code"),
            village_codes=data.get("village_codes"),
            hotspot_id=data.get("hotspot_id"),
            hotspot_code=data.get("hotspot_code"),
            catchment_id=data.get("catchment_id"),
            catchment_code=data.get("catchment_code"),
            exclude_verified_after=data.get("exclude_verified_after"),
            target_count=data.get("target_count"),
            benefit_plan_code=data.get("benefit_plan_code"),
        )
        projects = HouseholdValidationProjectLookupService().list_projects(
            location_id=(
                data.get("village_id")
                or data.get("ta_id")
                or data.get("district_id")
                or data.get("region_id")
            ),
            location_code=(
                data.get("village_code")
                or next(iter(data.get("village_codes") or []), None)
                or next(iter(data.get("gvh_codes") or []), None)
                or next(iter(data.get("ta_codes") or []), None)
                or data.get("ta_code")
                or data.get("district_code")
                or data.get("region_code")
            ),
            hotspot_id=data.get("hotspot_id"),
            hotspot_code=data.get("hotspot_code"),
            catchment_id=data.get("catchment_id"),
        )
        resolved_rule = service.eligibility_rule or {}
        quota_config = resolved_rule.get("selection_strategy")
        quota_percentages = (
            {
                "female_headed_percentage": female_headed_percentage(quota_config),
                "youth_headed_percentage": youth_headed_percentage(quota_config),
                "reserve_percentage": reserve_percentage(quota_config),
            }
            if quota_config is not None
            else {
                "female_headed_percentage": None,
                "youth_headed_percentage": None,
                "reserve_percentage": None,
            }
        )
        batch = HouseholdValidationBatch(
            district_id=data.get("district_id"),
            ta_id=data.get("ta_id"),
            village_id=data.get("village_id"),
            hotspot_code=data.get("hotspot_code"),
            catchment_code=data.get("catchment_code"),
            exclude_verified_after=data.get("exclude_verified_after"),
            target_count=data.get("target_count"),
            generated_at=timezone.now(),
            status=HouseholdValidationBatch.Status.PENDING,
            json_ext={
                "hotspot_id": data.get("hotspot_id"),
                "catchment_id": data.get("catchment_id"),
                "region_id": data.get("region_id"),
                "region_code": data.get("region_code"),
                "ta_codes": data.get("ta_codes") or [],
                "gvh_codes": data.get("gvh_codes") or [],
                "benefit_plan_code": data.get("benefit_plan_code"),
                **quota_percentages,
                "member_rows": len(selection_result.member_rows),
                **_json_safe(summary),
            },
        )
        batch.save(user=info.context.user)

        workbook_bytes = ExcelValidationListExporter(
            selection_result,
            batch_id=batch.id,
            projects=projects,
        ).export_bytes()
        file_name = f"household_validation_{batch.id}.xlsx"
        return HouseholdValidationGenerateResultGQLType(
            batch_id=batch.id,
            file_name=file_name,
            file_base64=base64.b64encode(workbook_bytes).decode("ascii"),
            **summary,
        )

    @staticmethod
    def _validate_user(user, perms):
        require_permissions(user, perms, error_class=PermissionDenied)


class UploadHouseholdValidationListMutation(graphene.Mutation):
    Output = HouseholdValidationUploadResultGQLType

    class Arguments:
        file_base64 = graphene.String(required=True)
        dry_run = graphene.Boolean(required=False)
        source_file_name = graphene.String(required=False)

    @classmethod
    def mutate(cls, root, info, **data):
        GenerateHouseholdValidationListMutation._validate_user(
            info.context.user,
            HouseholdValidationConfig.gql_mutation_upload_household_validation_list_perms,
        )
        try:
            workbook_bytes = base64.b64decode(data["file_base64"], validate=True)
        except Exception as exc:
            raise ValidationError("household_validation.upload.invalid_base64") from exc

        totals = HouseholdValidationUploadService(info.context.user).upload(
            workbook_bytes,
            dry_run=data.get("dry_run") or False,
            source_file_name=data.get("source_file_name"),
        )
        return HouseholdValidationUploadResultGQLType(**totals)
