import logging

from django.apps import AppConfig

logger = logging.getLogger(__name__)

MODULE_NAME = "household_validation"

# Config keys that used to be flat HouseholdValidationConfig class attributes
# before PWP's quota percentages moved into
# program_eligibility_rules["PWP"]["selection_strategy"]. Kept only so
# _load_config can warn a deployment whose saved ModuleConfiguration still
# sets them that the override no longer does anything.
RETIRED_CONFIG_KEYS = (
    "female_headed_percentage",
    "youth_percentage",
    "reserve_percentage",
)

DEFAULT_BUSINESS_TYPE_OPTIONS = [
    "Crop farming",
    "Livestock farming",
    "Grocery shop",
    "Tailoring",
    "Transport services",
    "Other businesses",
]

RIGHT_HOUSEHOLD_VALIDATION_QUERY_EXPORT = 958001
RIGHT_HOUSEHOLD_VALIDATION_UPLOAD = 958002
RIGHT_HOUSEHOLD_VALIDATION_HISTORY = 958003
RIGHT_HOUSEHOLD_VALIDATION_ERROR_REPORT = 958004

RIGHT_GROUP_SEARCH = 180001
RIGHT_GROUP_CREATE = 180002
RIGHT_GROUP_UPDATE = 180003
RIGHT_GROUP_DELETE = 180004

ROLE_DISTRICT_ADMINISTRATOR = "District Administrator"
ROLE_DISTRICT_PROGRAM_MANAGER = "District Program Manager"
ROLE_DISTRICT_USER = "District User"

HOUSEHOLD_VALIDATION_RIGHTS = [
    RIGHT_HOUSEHOLD_VALIDATION_QUERY_EXPORT,
    RIGHT_HOUSEHOLD_VALIDATION_UPLOAD,
    RIGHT_HOUSEHOLD_VALIDATION_HISTORY,
    RIGHT_HOUSEHOLD_VALIDATION_ERROR_REPORT,
]

GROUP_RIGHTS = [
    RIGHT_GROUP_SEARCH,
    RIGHT_GROUP_CREATE,
    RIGHT_GROUP_UPDATE,
    RIGHT_GROUP_DELETE,
]

DISTRICT_VALIDATION_ROLES = [
    ROLE_DISTRICT_ADMINISTRATOR,
    ROLE_DISTRICT_PROGRAM_MANAGER,
    ROLE_DISTRICT_USER,
]

DISTRICT_VALIDATION_ROLE_RIGHTS = {
    ROLE_DISTRICT_ADMINISTRATOR: HOUSEHOLD_VALIDATION_RIGHTS + [
        RIGHT_GROUP_SEARCH,
        RIGHT_GROUP_UPDATE,
    ],
    ROLE_DISTRICT_PROGRAM_MANAGER: HOUSEHOLD_VALIDATION_RIGHTS + [
        RIGHT_GROUP_SEARCH,
        RIGHT_GROUP_UPDATE,
    ],
    ROLE_DISTRICT_USER: HOUSEHOLD_VALIDATION_RIGHTS + [
        RIGHT_GROUP_SEARCH,
        RIGHT_GROUP_UPDATE,
    ],
}

DEFAULT_CONFIG = {
    "gql_query_household_validation_rule_perms": [
        str(RIGHT_HOUSEHOLD_VALIDATION_QUERY_EXPORT),
    ],
    "gql_mutation_generate_household_validation_list_perms": [
        str(RIGHT_HOUSEHOLD_VALIDATION_QUERY_EXPORT),
    ],
    "gql_mutation_upload_household_validation_list_perms": [
        str(RIGHT_HOUSEHOLD_VALIDATION_UPLOAD),
    ],
    "gql_query_household_validation_history_perms": [
        str(RIGHT_HOUSEHOLD_VALIDATION_HISTORY),
    ],
    "gql_query_household_validation_error_report_perms": [
        str(RIGHT_HOUSEHOLD_VALIDATION_ERROR_REPORT),
    ],
    "group_search_perms": [str(RIGHT_GROUP_SEARCH)],
    "group_update_perms": [str(RIGHT_GROUP_UPDATE)],
    "business_columns_enabled": False,
    "business_type_options": DEFAULT_BUSINESS_TYPE_OPTIONS,
    # Program Specific Eligibility + selection-strategy rules.
    #
    # - selection_strategy: presence/absence picks the algorithm. Omitted
    #   (the default), all eligible households are selected, up to
    #   target_count.
    # - requires_data_source: a member must have this value in their
    #   Individual.json_ext "data_source" key.
    # - requires_recipient_type: a member's GroupIndividual.recipient_type
    #   must match this value.
    # - member_flag: a member must have this Individual.json_ext boolean
    #   key to count towards the household's eligible members (a hard
    #   requirement, unlike priority_flag below).
    # - member_min_age / member_max_age: inclusive age bounds a member must
    #   also fall within to be eligible.
    # - priority_flag: a boolean key that ranks eligible households ahead
    #   of the rest of the pool when capping at target_count.
    "program_eligibility_rules": {
        "PWP": {
            "member_flag": "fit_for_work",
            "selection_strategy": {
                "female_headed_percentage": 40,
                "youth_headed_percentage": 40,
                "reserve_percentage": 20,
                "allocate_by_village": True,
            },
        },
        "RMEP": {
            "requires_data_source": "PWP",
            "priority_flag": "business_experience",
        },
        "UPG": {
            "requires_data_source": "SCTP",
            "member_flag": "fit_for_work",
            "member_min_age": 18,
            "member_max_age": 60,
        },
    },
}


class HouseholdValidationConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = MODULE_NAME

    gql_query_household_validation_rule_perms = None
    gql_mutation_generate_household_validation_list_perms = None
    gql_mutation_upload_household_validation_list_perms = None
    gql_query_household_validation_history_perms = None
    gql_query_household_validation_error_report_perms = None
    group_search_perms = None
    group_update_perms = None
    business_columns_enabled = DEFAULT_CONFIG["business_columns_enabled"]
    business_type_options = None
    program_eligibility_rules = None

    @classmethod
    def _load_config(cls, cfg):
        """
        Load config fields that match current AppConfig class fields.
        """
        # Existing deployments may have configuration saved before this flag.
        cls.business_columns_enabled = DEFAULT_CONFIG["business_columns_enabled"]
        stale_keys = [key for key in RETIRED_CONFIG_KEYS if key in cfg]
        if stale_keys:
            logger.warning(
                "household_validation ModuleConfiguration still sets %s, which no "
                "longer has any effect on selection -- these moved into "
                "program_eligibility_rules[\"PWP\"][\"selection_strategy\"]. Update "
                "this deployment's ModuleConfiguration to migrate the override, or "
                "it will silently use the default 40/40/20 quota split instead.",
                ", ".join(stale_keys),
            )
        for field in cfg:
            if hasattr(cls, field):
                setattr(cls, field, cfg[field])

    def ready(self):
        from core.models import ModuleConfiguration
        cfg = ModuleConfiguration.get_or_default(self.name, DEFAULT_CONFIG)
        self._load_config(cfg)
