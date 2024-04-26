import openpyxl
from openpyxl import Workbook
from utility import column_headers
import copy

# Sure wishing I had fuzzy search right about now to do this automatically from the sheet
betatel_main = ['ID', 'Data Type', 'victimdate', 'victimtime', 'Duration', 'endtime', 'A or B']
betatel_A = ['Terminating Number', 'Originating Party Start Location Name',
              'Terminating Party Start Location Name', 'Originating Party End Location Name', 'Terminating Party End Location Name', 'Originating Party Start Base Station ID', 'Terminating Party Start Base Station ID', 'Originating Party End Base Station ID', 'Terminating Party End Base Station ID',
              'Originating Party Start Location Latitude', 'Originating Party Start Location Longitude', 'Originating Party Start Location Bearing', 'Terminating Party Start Location Latitude', 'Terminating Party Start Location Longitude', 'Terminating Party Start Location Bearing',
              'Originating Party End Location Latitude', 'Originating Party End Location Longitude', 'Originating Party End Location Bearing', 'Terminating Party End Location Latitude', 'Terminating Party End Location Longitude', 'Terminating Party End Location Bearing']

# Create betatel_B automatically since I am not typing all that out again
betatel_B = copy.deepcopy(betatel_A)
for i in range(0, len(betatel_A)):
    if betatel_B[i][0] == 'T':
        betatel_B[i] = betatel_B[i].replace('Terminating', 'Originating', 1)
    elif betatel_B[i][0] == 'O':
        betatel_B[i] = betatel_B[i].replace('Originating', 'Terminating', 1)

# TODO Create functions for beta_extras to call to calculate victim date, time, and end time.


def beta_sheet(original_sheet, targetnumber, format):
    original_headers = column_headers(original_sheet)
    # Create result workbook and sheet
    wb_result = Workbook()
    ws_result = wb_result.active
    # Create headers
    for cell in format:
        ws_result.cell(row=1, column=format.index(cell)+1, value=cell)
    # Create data into every row, one row at a time
    for i in range(2, original_sheet.max_row+1):
        column = 0
        # Create main data
        for cell in betatel_main:
            column += 1
            if cell == 'A or B':
                continue
            if cell in original_headers:
                ws_result.cell(row=i, column=column, value=original_sheet.cell(row=i, column=original_headers[cell]).value)
            elif cell == 'ID':
                ws_result.cell(row=i, column=column, value=i - 1)
            else:  # expand with more elifs for specific cell types
                ws_result.cell(row=i, column=column, value='0')
                # Iterate over data that requires knowledge of whether A or B party.
        # If A party is our target, do this.
        if original_sheet.cell(row=i, column=original_headers['Originating Number']).value == targetnumber:
            # Set target to A party in sheet
            ws_result.cell(row=i, column=column, value='A')
            # Iterate over the party-dependent data and write to sheet
            for cell in betatel_A:
                column += 1
                if cell in original_headers:
                    ws_result.cell(row=i, column=column, value=original_sheet.cell(row=i, column=original_headers[cell]).value)
        # If A party is not our target, do this
        else:
            # Set target to B party in sheet
            ws_result.cell(row=i, column=column, value='B')
            # Iterate over party-dependent data but with reversed A and B party
            for cell in betatel_B:
                column += 1
                if cell in original_headers:
                    ws_result.cell(row=i, column=column, value=original_sheet.cell(row=i, column=original_headers[cell]).value)
    return wb_result
