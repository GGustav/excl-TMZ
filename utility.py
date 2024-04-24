def column_headers(worksheet):
    sheetcols = {}
    for i in range(1, worksheet.max_column + 1):
        sheetcols[worksheet.cell(row=1, column=i).value] = i
    return sheetcols