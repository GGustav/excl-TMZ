import os
import shutil

import openpyxl

from alphatel import alpha_sheet
from betatel import beta_sheet
from gammatel import gamma_sheet
from utility import target_providers, target_numbers

# Iterate over files in /originals
for file in os.listdir('originals'):
    target = os.path.splitext(file)[0]
    if target[0] == '~':
        continue
    origpath = "originals/" + file
    outputpath = "output/o_" + file
    # Open original workbook and sheet
    shutil.copyfile(origpath, outputpath)
    wb_orig = openpyxl.load_workbook(outputpath)

    # Use function based on provider
    match target_providers[target]:
        case 'AlphaTel':
            result = alpha_sheet(wb_orig, target_numbers[target])
        case 'BetaTel':
            result = beta_sheet(wb_orig, target_numbers[target])
        case 'GammaTel':
            result = gamma_sheet(wb_orig, target_numbers[target])

    result.save(filename=outputpath)
    print("Creation of", outputpath, "successful")
