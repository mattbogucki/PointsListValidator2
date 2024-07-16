import datetime


def version_is_expired() -> bool:

    expiration_date = datetime.datetime.strptime("11/1/2024", "%m/%d/%Y")
    todays_date = datetime.datetime.now()

    return todays_date > expiration_date
