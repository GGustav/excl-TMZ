from datetime import datetime, timedelta
import pytz
import timezonefinder
def column_headers(worksheet):
    sheetcols = {}
    for i in range(1, worksheet.max_column + 1):
        sheetcols[worksheet.cell(row=1, column=i).value] = i
    return sheetcols

# Target phone numbers
target_numbers = {'Jupiter' : '6148932837',
           'Mars' : '6141435201',
           'Mercury' : '6141928499',
           'Sol' : '6142358191',
           'Venus' : '6142819374'}

# Target providers
target_providers= {'Jupiter' : 'GammaTel',
             'Mars' : 'BetaTel',
             'Mercury' : 'BetaTel',
             'Sol' : 'AlphaTel',
             'Venus': 'AlphaTel'}

main_tz= pytz.timezone("Australia/Adelaide")

#Set format
format = ['ID', 'Event', 'Victim datetime', 'Target datetime', 'Target timezone', 'Duration', 'End time', 'A or B', 'Other number', 'Target start loc', 'Other start loc', 'Target end loc', 'Other end loc',
          'Target start base station', 'Other start base station', 'Target end base station', 'Other end base station', 'Target start lat', 'Target start lon', 'Target start bearing',
          'Other start lat', 'Other start lon', 'Other start bearing', 'Target end lat', 'Target end lon', 'Target end bearing', 'Other end lat', 'Other end lon', 'Other end bearing']

#Starting headers
main_headers = ['ID', 'Data Type', 'victimdatetime', 'targetdatetime', 'targettz', 'Duration', 'endtime', 'A or B']

def timezone_coord(lat, lon):
    tf = timezonefinder.TimezoneFinder()
    timezone_str = tf.certain_timezone_at(lat=lat, lng=lon)
    return timezone_str