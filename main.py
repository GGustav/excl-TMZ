import openpyxl
from openpyxl import Workbook
from alphatel import alpha_sheet
from betatel import beta_sheet
from gammatel import gamma_sheet
from utility import target_providers, target_numbers
import os

# Victim timezone - ACST. UTC+10:30 during winter, UTC+9:30 during summer.
# Switch to +9:30 happened Sun, Apr 2 at 3:00 AM local time

# Final format 1. All times in victim timezone
# ID - Event - Date - Time - Duration - End time - A or B party - Other number - Target start location - Other start location
# Target end loc - Other end loc - Target start base station ID - Other start base station ID - Target end base station ID - Other end base station ID -
# Target start lat - Target start lon - Target start bearing - Other start lat - Other start lon - Other start bearing
# Target end lat - Target end lon - Target end bearing - Other end lat - Other end lon - Other end bearing

# Maybe we add a 2nd worksheet with different formatting later in May?


# Create dictionary for column title as key and its position in the sheet as value



# Iterate over files in /originals
for file in os.listdir('originals'):
    target = os.path.splitext(file)[0]
    if target[0] == '~':
        continue
    origpath = "originals/" + file
    outputpath = "output/o_" + file
    # Open original workbook and sheet
    wb_orig = openpyxl.load_workbook(origpath)
    ws_orig = wb_orig.active

    # Use function based on provider
    match target_providers[target]:
        case 'AlphaTel':
            result = alpha_sheet(ws_orig, target_numbers[target])
        case 'BetaTel':
            result = beta_sheet(ws_orig, target_numbers[target])
        case 'GammaTel':
            result = gamma_sheet(ws_orig, target_numbers[target])

    result.save(filename=outputpath)
    print("Creation of", outputpath, "successful")
