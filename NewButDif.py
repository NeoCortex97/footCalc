from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('data/Simon_Schmlez_-_Basketball_SLJ_SS_BF_-_21-11-2023_-_Center_of_Force_line.xlsx')
ws = wb.active
# print(ws)

start_list = []
end_list = []
active = False

for cell in ws['C']:
   print(cell.value)
   if cell.value == "X (mm)":
      start_list.append(cell.row)
      active = True
   if active and cell.value is None:
      end_list.append(cell.row-1)
      active = False

messungen = str(len(start_list))
print("\nAnzahl der Messungen: " + messungen+"\n")
print(start_list)
print(end_list)

