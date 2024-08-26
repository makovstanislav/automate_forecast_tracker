import openpyxl
import xlwings as xw

# Objective 1: Скопировать из source в conversion

# Загрузка исходного Excel-файла source.xlsx
source = openpyxl.load_workbook('source.xlsx')
source_sheet = source.active

# Загрузка файла conversion.xlsx
conversion = openpyxl.load_workbook('conversion.xlsx')
conversion_sheet = conversion.worksheets[0]

# Копирование диапазона E5:G8 из source.xlsx в первый лист conversion.xlsx
source_range = source_sheet['E5:G8']
for i, row in enumerate(source_range):
    for j, cell in enumerate(row):
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            conversion_sheet.cell(row=5 + i, column=5 + j, value=cell.value)

# Сохранение изменений в файле conversion.xlsx
conversion.save('conversion.xlsx')

print("Objective 1 ✅")

# Objective 2: Копирование значений со второго листа conversion.xlsx в database.xlsx

# Загрузка conversion.xlsx с вычисленными значениями (data_only=True)
conversion_values = openpyxl.load_workbook('conversion.xlsx', keep_vba=True, data_only=True)
conversion_transposed = conversion_values.worksheets[1]  # Второй лист

wbxl=xw.Book('conversion.xlsx')

print(wbxl.sheets['Лист2'].range('D7').value)

# Загрузка файла database.xlsx
database = openpyxl.load_workbook('database.xlsx')
database_sheet = database.active

# Определение диапазона для копирования значений
transposed_range = conversion_transposed['C6:F8']

# Поиск первой пустой строки в database.xlsx
empty_row = database_sheet.max_row + 1

# Копирование значений (без формул) в database.xlsx
for i, row in enumerate(transposed_range):
    for j, cell in enumerate(row):
        database_sheet.cell(row=empty_row + i, column=3 + j, value=cell.value)

# Сохранение изменений в файле database.xlsx
database.save('database.xlsx')

print("Objective 2 ✅")