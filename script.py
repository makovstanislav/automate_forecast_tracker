import openpyxl

# Objective 1: Скопировать из source в output

# Get source file
source = openpyxl.load_workbook('source.xlsx')
source_sheet = source.active

# Get conversion file
conversion = openpyxl.load_workbook('output.xlsx')
conversion_sheet = conversion.worksheets[0]

source_range = source_sheet['E5:G8']

for i, row in enumerate(source_range):
    for j, cell in enumerate(row):
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            conversion_sheet.cell(row=5 + i, column=5 + j, value=cell.value)

conversion.save('output.xlsx')

print("Objective 1 ✅")

# Objective 2
conversion_transposed = conversion.worksheets[1]

database = openpyxl.load_workbook('database.xlsx')
database_sheet = database.active

transposed_range = conversion_transposed['C6:F8']

empty_row = database_sheet.max_row + 1

for i, row in enumerate(transposed_range):
    for j, cell in enumerate(row):
        database_sheet.cell(row=empty_row + i, column=3 + j, value=cell.value)

database.save('database.xlsx')

print("Objective 2 ✅")
