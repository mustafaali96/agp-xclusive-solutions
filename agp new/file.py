import pyodbc
import os
from tkcalendar import Calendar, DateEntry
import datetime
try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    import Tkinter as tk
    import ttk

def from_date():
    def print_sel():
        global start_date
        start_date = cal.selection_get()
        print("starting date is ", start_date)
        top.destroy()

    top = tk.Toplevel(root)
    today = datetime.date.today()
    maxdate = today + datetime.timedelta(days=0)
    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                   maxdate= maxdate, disabledforeground='red', cursor="hand1", year=2020, month=1, day=1)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()

def to_date():
    def print_sel():
        global end_date
        end_date = cal.selection_get()
        print("end date is ", end_date)
        top.destroy()


    top = tk.Toplevel(root)
    today = datetime.date.today()
    mindate = start_date
    maxdate = today + datetime.timedelta(days=0)

    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                    mindate=mindate, maxdate=maxdate, disabledforeground='red', cursor="hand1" )
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()


root = tk.Tk()
ttk.Button(root, text='FROM', command=from_date).pack(padx=10, pady=10)
ttk.Button(root, text='TO', command=to_date).pack(padx=10, pady=10)

root.mainloop()
try:
    os.remove("list.dat")  
except:
    pass

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./att2000.mdb;')
cursor = conn.cursor()
cursor.execute('select ui.Badgenumber , CHECKTIME , CHECKTYPE from CHECKINOUT as ch left join USERINFO as ui on ch.USERID = ui.USERID')
user_id= ""
date = ""
code = ""
for row in cursor.fetchall():
    user_id = row[0]
    while not (len(user_id) == 5):
        user_id = "0" + user_id    
    date = str(row[1]).replace("-","").replace(" ","").replace(":","")
    date = date[:-2]
    if(row[2] == "I"):
        code = "0001"
    elif(row[2]=="O"):
        code = "0002"
    line = "31" + date + code + "00000" + user_id + code
    with open("list.dat", "a") as my_file:
        my_file.write(line+"\n")
my_file.close