from datetime import datetime

import pytz

a = "2023-03-30 4:38:50"

date = datetime.strptime(a, "%Y-%m-%d %H:%M:%S")
utc = pytz.timezone("Etc/UTC")
adelaide = pytz.timezone("Australia/Adelaide")

date_tz = utc.localize(date)
localdate_tz = date_tz.astimezone(adelaide)

print(date_tz)
print(localdate_tz)

lat = -35.13993
lon = 138.47236
