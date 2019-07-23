from xlrd import open_workbook
from datetime import date
from xlutils.copy import  copy

wb = open_workbook('Book1.xlsx')
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

read_result = open_workbook('result.xls')
curr_sheet = read_result.sheet_by_index(0)

last_row = curr_sheet.nrows
i = last_row
if last_row != 0:
    last_row += 1

res=copy(read_result)
result=res.get_sheet(0)

reciept = -1
if i == 0:
    reciept = 1
else:
    while curr_sheet.cell_value(i-1,0) != 'reciept no' and i > 1:
        x = curr_sheet.cell_value(i-1, 0)
        i -= 1
    reciept = curr_sheet.cell_value(i-1, 1)+1
result.write(last_row,0,'reciept no')
result.write(last_row, 1, reciept)

last_row+=1
result.write(last_row, 0, 'name :')
print('Enter name')
cust_name = input()
result.write(last_row, 1, cust_name)

last_row+=1
result.write(last_row,0,'date :')
today = date.today()
result.write(last_row, 1, today)


last_row+=1
result.write(last_row, 0, 'item')
result.write(last_row, 1, 'price')
result.write(last_row, 2, 'quantity')
result.write(last_row, 3, 'total')

last_row += 1
print("Enter total no of items")
t = int(input())
total = 0
no = []
q = []
print('Enter item nos and respective quantity')
while t > 0:
    t -= 1
    a, b = map(int, input().split())
    no.append(a)
    q.append(b)
print('::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::')
for i in range(0, len(no)):
    total += q[i] * sheet.cell_value(no[i], 2)
    item=sheet.cell_value(no[i], 1)
    price=sheet.cell_value(no[i], 2)
    print(item, '\t\t\tprice', price, " x", q[i], ' \t:\t ',
          q[i] * price)
    result.write(last_row,0,item)
    result.write(last_row,1,price)
    result.write(last_row, 2, q[i])
    result.write(last_row, 3,q[i]*price)
    last_row+=1

print('total bill:\t\t\t\t\t\t\t\t\t\t\t', total)
result.write(last_row,2,'total bill')
result.write(last_row,3 , total)
res.save('result.xls')
