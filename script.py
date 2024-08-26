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

# Objective 2: Использование xlwings для расчета и получения значений

# Открываем файл с помощью xlwings
wbxl = xw.Book('conversion.xlsx')

# Выбираем второй лист
sheet2 = wbxl.sheets['Лист2']

# Извлечение значений из диапазона C6:F8 после автоматического пересчета формул
calculated_values = sheet2.range('C6:F8').value

# Закрытие файла, сохранение не нужно, так как значения уже извлечены
wbxl.close()

# Загрузка файла database.xlsx
database = openpyxl.load_workbook('database.xlsx')
database_sheet = database.active

# Поиск первой пустой строки в database.xlsx
empty_row = database_sheet.max_row + 1

# Вставка вычисленных значений в database.xlsx
for i, row in enumerate(calculated_values):
    for j, value in enumerate(row):
        database_sheet.cell(row=empty_row + i, column=3 + j, value=value)

# Сохранение изменений в файле database.xlsx
database.save('database.xlsx')

print("Objective 2 ✅")