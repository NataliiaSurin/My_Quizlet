import openpyxl as op
import random

df=op.load_workbook('Courses.xlsx')
data=df.active
total_rows=data.max_row
#columns=data.max_column
print(total_rows)
words=[]
definitions=[]

#Creating list with words
for row_number in range(3,total_rows+1):
    cell_obj = data.cell(row=row_number, column=1)
    words.append(cell_obj.value)
print(words)

#Creating list with definitions
for row_number in range(3,total_rows+1):
    cell_obj = data.cell(row=row_number, column=3)
    definitions.append(cell_obj.value)
print(definitions)

random_row_number=int(random.uniform(3,total_rows+1))
print("For word *",words[random_row_number],"* definition is *",definitions[random_row_number],"*")

