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