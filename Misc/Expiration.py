import datetime


def version_is_expired() -> bool:

    expiration_date = datetime.datetime.strptime("4/1/2025", "%m/%d/%Y")
    todays_date = datetime.datetime.now()

    return todays_date > expiration_date
