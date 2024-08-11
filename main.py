import openpyxl as op
df=op.load_workbook('Courses.xlsx')
data=df.active
total_rows=data.max_row
#columns=data.max_column
print(total_rows)
for row_number in range(3,total_rows+1):
    cell_obj = data.cell(row=row_number, column=1)
    word=cell_obj.value
    print(word)

