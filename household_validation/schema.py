import base64

import graphene
from django.core.exceptions import PermissionDenied

from household_validation.apps import HouseholdValidationConfig
from household_validation.gql_permissions import require_permissions
from household_validation.excel import (
    build_rejected_households_workbook_bytes,
    is_primary_worker_rejection,
)
from household_validation.gql_mutations import (
    GenerateHouseholdValidationListMutation,
    UploadHouseholdValidationListMutation,
)
from household_validation.gql_queries import (
    HouseholdValidationBatchRowsGQLType,
    HouseholdValidationBatchesGQLType,
    HouseholdValidationErrorReportGQLType,
    HouseholdValidationPreviewEdgeGQLType,
    HouseholdValidationPreviewGQLType,
    HouseholdValidationPreviewPageInfoGQLType,
    HouseholdValidationProjectsGQLType,
)
from household_validation.models import (
    HouseholdValidationBatch,
    HouseholdValidationBatchRow,
)
from household_validation.services import (
    EligibleHouseholdSelectionService,
    HouseholdValidationProjectLookupService,
    build_validation_error_report,
)

VALIDATION_FILTER_ARG_NAMES = {
    "region_id",
    "region_code",
    "district_id",
    "district_code",
    "ta_id",
    "ta_code",
    "ta_codes",
    "gvh_codes",
    "village_id",
    "village_code",
    "village_codes",
    "hotspot_id",
    "hotspot_code",
    "catchment_id",
    "catchment_code",
    "exclude_verified_after",
    "target_count",
    "benefit_plan_code",
}


def validation_filter_args():
    return {
        "region_id": graphene.Argument(graphene.Int, required=False),
        "region_code": graphene.Argument(graphene.String, required=False),
        "district_id": graphene.Argument(graphene.Int, required=False),
        "district_code": graphene.Argument(graphene.String, required=False),
        "ta_id": graphene.Argument(graphene.Int, required=False),
        "ta_code": graphene.Argument(graphene.String, required=False),
        "ta_codes": graphene.Argument(graphene.List(graphene.String), required=False),
        "gvh_codes": graphene.Argument(graphene.List(graphene.String), required=False),
        "village_id": graphene.Argument(graphene.Int, required=False),
        "village_code": graphene.Argument(graphene.String, required=False),
        "village_codes": graphene.Argument(graphene.List(graphene.String), required=False),
        "hotspot_id": graphene.Argument(graphene.String, required=False),
        "hotspot_code": graphene.Argument(graphene.String, required=False),
        "catchment_id": graphene.Argument(graphene.String, required=False),
        "catchment_code": graphene.Argument(graphene.String, required=False),
        "exclude_verified_after": graphene.Argument(graphene.Date, required=False),
        "target_count": graphene.Argument(graphene.Int, required=False),
        "benefit_plan_code": graphene.Argument(graphene.String, required=False),
    }


class Query(graphene.ObjectType):
    household_validation_preview = graphene.Field(
        HouseholdValidationPreviewGQLType,
        first=graphene.Argument(graphene.Int, required=False),
        offset=graphene.Argument(graphene.Int, required=False),
        order_by=graphene.Argument(graphene.String, required=False),
        **validation_filter_args(),
    )
    household_validation_projects = graphene.Field(
        HouseholdValidationProjectsGQLType,
        location_id=graphene.Argument(graphene.Int, required=False),
        location_code=graphene.Argument(graphene.String, required=False),
        hotspot_id=graphene.Argument(graphene.String, required=False),
        hotspot_code=graphene.Argument(graphene.String, required=False),
        catchment_id=graphene.Argument(graphene.String, required=False),
    )
    household_validation_batches = graphene.Field(
        HouseholdValidationBatchesGQLType,
        status=graphene.Argument(graphene.String, required=False),
    )
    household_validation_batch_rows = graphene.Field(
        HouseholdValidationBatchRowsGQLType,
        batch_id=graphene.Argument(graphene.UUID, required=True),
        status=graphene.Argument(graphene.String, required=False),
    )
    household_validation_rejected_batch_rows = graphene.Field(
        HouseholdValidationBatchRowsGQLType,
        batch_id=graphene.Argument(graphene.UUID, required=True),
        upload_attempt_id=graphene.Argument(graphene.UUID, required=True),
    )
    household_validation_batch_error_report = graphene.Field(
        HouseholdValidationErrorReportGQLType,
        batch_id=graphene.Argument(graphene.UUID, required=True),
    )

    def resolve_household_validation_preview(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_query_household_validation_rule_perms,
        )
        rows = EligibleHouseholdSelectionService(info.context.user).preview(
            **Query._selection_filters(kwargs)
        )
        total_count = len(rows)
        offset = max(0, kwargs.get("offset") or 0)
        first = kwargs.get("first")
        if first is None:
            first = total_count
        first = max(0, first)
        page_rows = rows[offset:offset + first]
        return HouseholdValidationPreviewGQLType(
            edges=[
                HouseholdValidationPreviewEdgeGQLType(node=row)
                for row in page_rows
            ],
            total_count=total_count,
            page_info=HouseholdValidationPreviewPageInfoGQLType(
                has_next_page=offset + first < total_count,
                has_previous_page=offset > 0,
                start_cursor=str(offset) if page_rows else None,
                end_cursor=str(offset + len(page_rows) - 1) if page_rows else None,
            ),
        )

    def resolve_household_validation_projects(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_query_household_validation_rule_perms,
        )
        projects = HouseholdValidationProjectLookupService().list_projects(
            location_id=kwargs.get("location_id"),
            location_code=kwargs.get("location_code"),
            hotspot_id=kwargs.get("hotspot_id"),
            hotspot_code=kwargs.get("hotspot_code"),
            catchment_id=kwargs.get("catchment_id"),
        )
        return HouseholdValidationProjectsGQLType(
            projects=projects,
            count=len(projects),
        )

    def resolve_household_validation_batches(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_query_household_validation_history_perms,
        )
        queryset = HouseholdValidationBatch.objects.filter(is_deleted=False)
        if kwargs.get("status"):
            queryset = queryset.filter(status=kwargs["status"])
        batches = list(queryset.order_by("-date_created"))
        return HouseholdValidationBatchesGQLType(
            batches=batches,
            count=len(batches),
        )

    def resolve_household_validation_batch_rows(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_query_household_validation_history_perms,
        )
        queryset = HouseholdValidationBatchRow.objects.filter(
            batch_id=kwargs["batch_id"],
            is_deleted=False,
        )
        if kwargs.get("status"):
            queryset = queryset.filter(status=kwargs["status"])
        rows = list(queryset.order_by("row_number"))
        return HouseholdValidationBatchRowsGQLType(
            rows=rows,
            count=len(rows),
        )

    def resolve_household_validation_rejected_batch_rows(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_mutation_upload_household_validation_list_perms,
        )
        rows = list(
            HouseholdValidationBatchRow.objects.filter(
                batch_id=kwargs["batch_id"],
                upload_attempt_id=kwargs["upload_attempt_id"],
                status=HouseholdValidationBatchRow.Status.REJECTED,
                is_deleted=False,
            ).order_by("row_number")
        )
        report, _ = build_rejected_households_workbook_bytes(rows)
        return HouseholdValidationBatchRowsGQLType(
            rows=rows,
            count=len(rows),
            file_name=(
                f"rejected_households_{kwargs['batch_id']}_"
                f"{kwargs['upload_attempt_id']}.xlsx"
            ),
            file_base64=base64.b64encode(report).decode("ascii"),
        )

    def resolve_household_validation_batch_error_report(parent, info, **kwargs):
        Query._check_permissions(
            info.context.user,
            HouseholdValidationConfig.gql_query_household_validation_error_report_perms,
        )
        batch = HouseholdValidationBatch.objects.get(
            id=kwargs["batch_id"],
            is_deleted=False,
        )
        error_rows = [
            row
            for row in batch.rows.filter(
                status=HouseholdValidationBatchRow.Status.ERROR,
                is_deleted=False,
            )
            if not is_primary_worker_rejection(row)
        ]
        report = build_validation_error_report(batch)
        return HouseholdValidationErrorReportGQLType(
            batch_id=batch.id,
            file_name=f"household_validation_errors_{batch.id}.csv",
            file_base64=base64.b64encode(report.encode("utf-8")).decode("ascii"),
            error_count=len(error_rows),
        )

    @staticmethod
    def _check_permissions(user, perms):
        require_permissions(user, perms, error_class=PermissionDenied)

    @staticmethod
    def _selection_filters(kwargs):
        return {
            key: kwargs.get(key)
            for key in VALIDATION_FILTER_ARG_NAMES
        }


class Mutation(graphene.ObjectType):
    generate_household_validation_list = GenerateHouseholdValidationListMutation.Field()
    upload_household_validation_list = UploadHouseholdValidationListMutation.Field()
