from django.apps import AppConfig

MODULE_NAME = "household_validation"

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
    "female_headed_percentage": 40,
    "youth_percentage": 40,
    "reserve_percentage": 20,
    "business_columns_enabled": False,
    "business_type_options": DEFAULT_BUSINESS_TYPE_OPTIONS,
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
    female_headed_percentage = None
    youth_percentage = None
    reserve_percentage = None
    business_columns_enabled = DEFAULT_CONFIG["business_columns_enabled"]
    business_type_options = None

    @classmethod
    def _load_config(cls, cfg):
        """
        Load config fields that match current AppConfig class fields.
        """
        # Existing deployments may have configuration saved before this flag.
        cls.business_columns_enabled = DEFAULT_CONFIG["business_columns_enabled"]
        for field in cfg:
            if hasattr(cls, field):
                setattr(cls, field, cfg[field])

    def ready(self):
        from core.models import ModuleConfiguration
        cfg = ModuleConfiguration.get_or_default(self.name, DEFAULT_CONFIG)
        self._load_config(cfg)
