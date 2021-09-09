import openpyxl
filename = "C:\\Users\Bill\VetsInTech\itp_week_3\day_1\lecture.xlsx"
wb = openpyxl.load_workbook(filename)
my_sheet = wb['Sheet1']

# for i in range(1, 8):
#     print(i, my_sheet.cell(row = i, column = 1).value)
#     print(i, my_sheet.cell(row = i, column = 2).value)
#     print(i, my_sheet.cell(row = i, column = 3).value)




row_count = my_sheet.max_row
column_count = my_sheet.max_column

row_count += 1
column_count += 1
print(row_count)
print(column_count)

for i in range(1, row_count):
    print(i, my_sheet.cell(row = i, column = 1).value)
    print(i, my_sheet.cell(row = i, column = 2).value)
    print(i, my_sheet.cell(row = i, column = 3).value)