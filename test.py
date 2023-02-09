import datetime

expiration_date = datetime.datetime.strptime("1/1/2023", "%m/%d/%Y")
todays_date = datetime.datetime.now()

print(todays_date > expiration_date)

print('None'.split("_")[-1])