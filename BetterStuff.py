import pandas as pd
from openpyxl import load_workbook
import numpy as np
import tkinter as tk
from tkinter import filedialog
path=""

def Fehlermeldung():
    Fehler = tk.Tk()
    Fehler.title("Fehler beim Öffnen der Datei")
    Fehler.geometry("500x100")
    Fehlertxt = tk.Label(Fehler, text="Bitte geben Sie den Dateityp .xlsx ein.")
    okbutt= tk.Button(Fehler, text ="OK", command=Fehler.destroy)
    Fehlertxt.pack()
    okbutt.pack()
    Fehler.mainloop()

def Hilfetext():
    textroot = tk.Tk()
    textroot.title("Hilfe")
    textroot.geometry("900x900")
    data = open('data/hilftxt', "r")
    Hilfetxt = tk.Label(textroot, text=data.read())
    beendenbutton = tk.Button(textroot, text="Schließen", command=textroot.destroy)
    Hilfetxt.pack(side="left")
    beendenbutton.pack(anchor='ne',side="top")
    textroot.mainloop()

#from openpyxl.utils import get_column_letter

def get_file_path():
    global path
    # Dateidialog öffnen, um einen Dateipfad auszuwählen
    path = filedialog.askopenfilename()

    # Den ausgewählten  Dateipfad anzeigen
    if path:
        path_label.config(text="Dateipfad: " + path)
    else:
        path_label.config(text="Keine Datei ausgewählt")


def Hauptfenster():
    root = tk.Tk()
    root.title("Dateipfad auswählen")
    root.state('zoomed')
    def ren ():
        if check_path() == True:
            root.destroy()
    # Button zum Öffnen des Dateidialogs
    browse_button = tk.Button(root, text="Datei auswählen", command=get_file_path)
    browse_button2 = tk.Button(root, text="Bestätigen", command= ren)
    browse_hilfe = tk.Button(root, text="Hilfe", command=Hilfetext, bg="#FF4500")
    browse_hilfe.place(relx=1.0, rely=0.0, anchor='ne', x=-10, y=10)
    browse_button.pack(anchor='n', pady=10)
    browse_button2.pack(pady=10)



    # Label zur Anzeige des ausgewählten Dateipfads
    global path_label
    path_label = tk.Label(root, text="Dateipfad: ")
    path_label.pack()


    root.mainloop()

def read_clean_val(messungen, start_list, end_list, Data):

   i = 0
   data = []
   while i < messungen:
      start = start_list[i]
      end = end_list[i]
      #print(Data_x[start:end])
      data.append(Data[start:end])
      i += 1
   #print(Data_x)
   return data


def get_range():
   start = []
   end = []
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


def get_data():
   Data_x = []
   Data_y = []
   for cell in ws['C']:
      Data_x.append(cell.value)
   for cell in ws['D']:
      Data_y.append(cell.value)
   return Data_x, Data_y

def get_df(data):
    h= 0
    minLen=len(data[0])
    for h in range(len(data)):
        if len(data[h])<=minLen:
            minLen=len(data[h])
        h+=1

    for k in range(len(data)):
        data[k]=np.array(data[k][:minLen])

    frame = []
    for i in range(minLen):
        frame.append(i)

    #print(frame)
    #print(minLen)

    df=pd.DataFrame(
        {
            'Frames': frame,
            'LinFus1': data[0],
            'LinFus2': data[1],
            'LinFus3': data[2],
            'RecFus1': data[3],
            'RecFus2': data[4],
            'RecFus3': data[5],
        }
    )
    return df

def Fus_mit(df1):
    mit_left_foot = []
    mit_right_foot =[]
    for c in range(df1.shape[0]):
        mit_left_foot.append(round(((df1.at[c, 'LinFus1'] + df1.at[c, 'LinFus2'] + df1.at[c, 'LinFus3'])/3), 2))
    for i in range(df1.shape[0]):
        mit_right_foot.append(round(((df1.at[i, 'RecFus1'] + df1.at[i, 'RecFus2'] + df1.at[i, 'RecFus3'])/3), 2))
    #print("TEst")
    frame = []
    for i in range(len(df1)):
        frame.append(i)
    sec_df = pd.DataFrame(
        {
            'kek': frame,
            'Mit_Left_Val': mit_left_foot,
            'Mit_Right_Val': mit_right_foot
        }
    )
    #print("2Test")
    dfMerge = df1.merge(sec_df, how='inner', right_on='kek', left_on='Frames')
    dfMerge.drop('kek',axis=1, inplace=True)
    #print(dfMerge)
    return dfMerge

def check_path():
    global path
    if ".xlsx" not in path:
        Fehlermeldung()
        return "NotX"
    else :
        return True

def auswerutng() :
    #if check_path() != "NotX":
      #  wb = load_workbook(path)
      #  ws = wb.active
        #pop Up fenster
        datenausertung = tk.Tk()
        datenausertung.title("Auswertung")
        datenausertung.state('zoomed')

        Data_x, Data_y = get_data()
        start_list, end_list = get_range()
        end_list.append(len(ws['c']))
        messungen = (len(start_list))
        print("Anzahl der Messungen: ", messungen)

        df_x = Fus_mit(get_df(read_clean_val(messungen, start_list, end_list, Data_x)))
        df_y = Fus_mit(get_df(read_clean_val(messungen, start_list, end_list, Data_y)))
        print(df_x)
        print(df_y)

        #ausgabe des data Frames
        daten_label = tk.Label(datenausertung, text=""+df_x.to_string())
        daten_label.pack()
        datenausertung.mainloop()





Hauptfenster()
wb = load_workbook(path)
ws = wb.active
auswerutng()





