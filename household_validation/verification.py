VERIFIED = "VERIFIED"
NOT_VERIFIED = "NOT_VERIFIED"
REJECTED = "REJECTED"

PARTICIPANT_STATUS_COLUMN = "participant_status"
HOUSEHOLD_STATUS_COLUMN = "household_status"
BUSINESS_REJECTION_CODE = "BUSINESS_WITHOUT_PRIMARY_WORKER"


def yes_no(value):
    if value is None:
        return None
    value = str(value).strip().upper()
    if value in {"YES", "Y", "TRUE", "1"}:
        return True
    if value in {"NO", "N", "FALSE", "0"}:
        return False
    return None


def participant_status(primary_worker, business, *, business_columns_enabled=True):
    if not business_columns_enabled:
        return VERIFIED if primary_worker is True else NOT_VERIFIED
    business = yes_no(business)
    if primary_worker is True and business is not None:
        return VERIFIED
    # Business information on a non-primary worker is an upload validation error.
    return NOT_VERIFIED


def household_status(rows, *, business_columns_enabled=True):
    """Reject conflicting households before considering a qualifying worker."""
    statuses = [
        participant_status(worker, business, business_columns_enabled=business_columns_enabled)
        for worker, business in rows
    ]
    if sum(worker is True for worker, _ in rows) > 1 or REJECTED in statuses:
        return REJECTED
    return VERIFIED if VERIFIED in statuses else NOT_VERIFIED
