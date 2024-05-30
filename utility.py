from datetime import datetime, timedelta
import pytz
import timezonefinder
import copy

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

#Main timezone we are looking for
main_tz= pytz.timezone("Australia/Adelaide")

#Set output format
format = ['ID', 'Event', 'Victim datetime', 'Target datetime', 'Target timezone', 'Duration', 'End time', 'A or B', 'Other number', 'Target start loc', 'Other start loc', 'Target end loc', 'Other end loc',
          'Target start base station', 'Other start base station', 'Target end base station', 'Other end base station', 'Target start lat', 'Target start lon', 'Target start bearing',
          'Other start lat', 'Other start lon', 'Other start bearing', 'Target end lat', 'Target end lon', 'Target end bearing', 'Other end lat', 'Other end lon', 'Other end bearing']

#Starting headers
main_headers = ['ID', 'Data Type', 'victimdatetime', 'targetdatetime', 'targettz', 'Duration', 'endtime', 'A or B']


#-- A start Latitude and longitude have to be 10th and 11th in the array, B start latitude 13 and longitude 14, and terminating/B number has to be first

#------------------------
#Alphatel format
alphatel_A = ['B Number', 'A Start Location Name',
              'B Start Location Name', 'A End Location Name', 'B End Location Name', 'A Start Location CGI', 'B Start Location CGI', 'A End Location CGI', 'B End Location CGI',
              'A Start Location Latitude', 'A Start Location Longitude', 'A Start Location Bearing', 'B Start Location Latitude', 'B Start Location Longitude', 'B Start Location Bearing',
              'A End Location Latitude', 'A End Location Longitude', 'A End Location Bearing', 'B End Location Latitude', 'B End Location Longitude', 'B End Location Bearing']

# Create alphatel_B automatically since I am not typing all that out again
alphatel_B = copy.deepcopy(alphatel_A)
for i in range(0, len(alphatel_A)):
    if alphatel_B[i][0] == 'A':
        alphatel_B[i] = alphatel_B[i].replace('A', 'B', 1)
    elif alphatel_B[i][0] == 'B':
        alphatel_B[i] = alphatel_B[i].replace('B', 'A', 1)

# Alphatel provided timezone
alpha_tz = pytz.timezone("Etc/UTC")

#----------------------------
#Betatel format
betatel_A = ['Terminating Number', 'Originating Party Start Location Name',
              'Terminating Party Start Location Name', 'Originating Party End Location Name', 'Terminating Party End Location Name', 'Originating Party Start Location CGI', 'Terminating Party Start Location CGI', 'Originating Party End Location CGI', 'Terminating Party End Location CGI',
              'Originating Party Start Location Latitude', 'Originating Party Start Location Longitude', 'Originating Party Start Location Bearing', 'Terminating Party Start Location Latitude', 'Terminating Party Start Location Longitude', 'Terminating Party Start Location Bearing',
              'Originating Party End Location Latitude', 'Originating Party End Location Longitude', 'Originating Party End Location Bearing', 'Terminating Party End Location Latitude', 'Terminating Party End Location Longitude', 'Terminating Party End Location Bearing']

# Create betatel_B automatically since I am not typing all that out again
betatel_B = copy.deepcopy(betatel_A)
for i in range(0, len(betatel_A)):
    if betatel_B[i][0] == 'T':
        betatel_B[i] = betatel_B[i].replace('Terminating', 'Originating', 1)
    elif betatel_B[i][0] == 'O':
        betatel_B[i] = betatel_B[i].replace('Originating', 'Terminating', 1)

# Betatel provided timezone
beta_tz = pytz.timezone("Australia/Perth")

#---------------------
#Gammatel format
gammatel_A = ['B Number', 'A Party Start Location Name',
              'B Party Start Location Name', 'A Party End Location Name', 'B Party End Location Name', 'A Party Start Base Station ID', 'B Party Start Base Station ID', 'A Party End Base Station ID', 'B Party End Base Station ID',
              'A Party Start Location Latitude', 'A Party Start Location Longitude', 'A Party Start Location Bearing', 'B Party Start Location Latitude', 'B Party Start Location Longitude', 'B Party Start Location Bearing',
              'A Party End Location Latitude', 'A Party End Location Longitude', 'A Party End Location Bearing', 'B Party End Location Latitude', 'B Party End Location Longitude', 'B Party End Location Bearing']

# Create gammatel_B automatically since I am not typing all that out again
gammatel_B = copy.deepcopy(gammatel_A)
gammatel_B[0] = 'A Party Number'
for i in range(1, len(gammatel_A)):
    if gammatel_B[i][0] == 'A':
        gammatel_B[i] = gammatel_B[i].replace('A', 'B', 1)
    elif gammatel_B[i][0] == 'B':
        gammatel_B[i] = gammatel_B[i].replace('B', 'A', 1)

# Gammatel provided timezone
gamma_tz = pytz.timezone("Etc/UTC")

def timezone_coord(lat, lon):
    tf = timezonefinder.TimezoneFinder()
    timezone_str = tf.certain_timezone_at(lat=lat, lng=lon)
    return timezone_str

def column_headers(worksheet):
    sheetcols = {}
    for i in range(1, worksheet.max_column + 1):
        sheetcols[worksheet.cell(row=1, column=i).value] = i
    return sheetcols
