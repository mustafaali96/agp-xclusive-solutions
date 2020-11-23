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

start_date = None
end_date = None

def from_date():
    def print_sel():
        global start_date
        start_date = cal.selection_get()
        print("starting date is ", start_date)
        ttk.Label(root, text="From "+str(start_date)).pack(padx=10, pady=10)
        top.destroy()

    top = tk.Toplevel(root)
    today = datetime.date.today()
    maxdate = today + datetime.timedelta(days=0)
    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                   maxdate= maxdate, disabledforeground='red', cursor="hand1", year=2020, month=1, day=1)
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()
    if end_date != None:
        ttk.Button(root, text="POOL",command=exitt).pack(padx=10, pady=10)
    else:
        pass

def to_date():
    def print_sel():
        global end_date
        end_date = cal.selection_get()
        print("end date is ", end_date)
        ttk.Label(root, text="To "+str(end_date)).pack(padx=10, pady=10)
        top.destroy()


    top = tk.Toplevel(root)
    today = datetime.date.today()
    mindate = start_date
    maxdate = today + datetime.timedelta(days=0)

    cal = Calendar(top, font="Arial 14", selectmode='day', locale='en_US',
                    mindate=mindate, maxdate=maxdate, disabledforeground='red', cursor="hand1" )
    cal.pack(fill="both", expand=True)
    ttk.Button(top, text="ok", command=print_sel).pack()
    if start_date != None:
        ttk.Button(root, text="POOL",command=exitt).pack(padx=10, pady=10)
    else:
        pass
    

def exitt():
    root.destroy()
    
root = tk.Tk()
root.title("AGP") 
root.geometry("300x300")

ttk.Button(root, text='FROM', command=from_date).pack(padx=10, pady=10)

ttk.Button(root, text='TO', command=to_date).pack(padx=10, pady=10)

root.mainloop()

try:
    os.remove("list.dat")  
except:
    pass
start_date = str(start_date).split("-")
from_date = "#" + start_date[1] + "/" + start_date [2] + "/" + start_date[0] + " 00:00:00#"

end_date = str(end_date).split("-")
to_date = "#" + end_date[1] + "/" + end_date [2] + "/" + end_date[0] + " 23:59:59#"


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=./att2000.mdb;')
cursor = conn.cursor()
sql = "SELECT ui.Badgenumber, ch.CHECKTIME, ch.CHECKTYPE FROM CHECKINOUT AS ch LEFT JOIN USERINFO AS ui ON ch.USERID = ui.USERID WHERE (((ch.CHECKTIME)BETWEEN %s AND %s)) order by ch.CHECKTIME desc"%(from_date,to_date)
#cursor.execute('select ui.Badgenumber , CHECKTIME , CHECKTYPE from CHECKINOUT as ch left join USERINFO as ui on ch.USERID = ui.USERID')
cursor.execute(sql)
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