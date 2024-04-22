from openpyxl import Workbook, load_workbook
#from openpyxl.utils import get_column_letter
import pandas as pd

def read_clean_val(messungen, start_list, end_list,Data):
   i = 0
   data = []
   while i < messungen:
      start = start_list[i]
      end = end_list[i]
      #print(Data_x[start:end])
      data.append(Data[start:end])
      i += 1
   #print(Data_x)
   return (data)

def get_range (start = [], end = []):
   active = False
   for cell in ws['A']:
      if cell.value == 'Frame':
         start.append(cell.row)
         active = True
      #print(cell.value)
      if active and cell.value is None:
         end.append(cell.row-1)
         active = False
   return (start, end)

def get_data ():
   for cell in ws['C']:
      Data_x.append(cell.value)
   for cell in ws['D']:
      Data_y.append(cell.value)
   return (Data_x, Data_y)


path = "data/Simon_Schmlez_-_Basketball_SLJ_SS_BF_-_21-11-2023_-_Center_of_Force_line.xlsx"
wb = load_workbook(path)
ws = wb.active
Data_x = []
Data_y = []
Data_x, Data_y = get_data()
start_list, end_list = get_range()

end_list.append(len(ws['c']))
messungen = (len(start_list))
print("Anzahl der Messungen: ", messungen)
print(start_list)
print(end_list)

data_x = read_clean_val(messungen, start_list, end_list, Data_x)
data_y = read_clean_val(messungen, start_list, end_list, Data_y)
#print(data_x)
#print(data_x[2])
#print(Data_y)
print(data_y[0])
print(data_x[0])
