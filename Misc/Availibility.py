from enum import Enum


class Availability(Enum):
    NOT_REQUESTED_AVAILABLE = "Not Requested-Available"
    REQUESTED_AVAILABLE = "Requested-Available"
    REQUESTED_NOT_APPLICABLE = "Requested-Not Applicable"
    REQUESTED_NOT_AVAILABLE = "Requested-Not Available"
