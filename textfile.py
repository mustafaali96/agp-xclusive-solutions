import xlrd
import os
try:
    os.remove("list.dat")
    
except:
    pass
workbook = xlrd.open_workbook('daily.xls')
worksheet = workbook.sheet_by_index(0)
first_row = [] # Header
for col in range(worksheet.ncols):
    first_row.append( worksheet.cell_value(0,col) )
# tronsform the workbook to a list of dictionnaries
data =[]
ids = ""
date_time = ""
status = ""
code = ""
line = ""
add_hours = 0

for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]]=worksheet.cell_value(row,col)
    for col_name, value in elm.items():
        if col_name == "No.":
            while not (len(value) == 5):
                value = "0" + value
            ids = value
            #print("No is: ", value)
        
        if col_name == "Date/Time":
            try:
                date_time = value.replace(":","").replace("-","").replace(".","").replace("Jan","0120").replace("Feb","0220").replace("Mar","0320").replace("Apr","0420").replace("May","0520").replace("Jun","0620").replace("Jul","0720").replace("Aug","0820").replace("Sep","0920").replace("Oct","1020").replace("Nov","1120").replace("Dec","1220")
            except:
                pass
            
            datetime_list = date_time.split(" ")
            date_month_year = datetime_list[0].split("/")
            
            if(len(date_month_year[0]) == 1):
                date_month_year[0] = "0" + date_month_year[0]
            else:
                pass
            if(len(date_month_year[1]) == 1):
                date_month_year[1] = "0" + date_month_year[1]
            else:
                pass
            
            date = date_month_year[1]
            month = date_month_year[0]
            year = date_month_year[2]
            hours = 0
            mins = 0
            if(datetime_list[2] == "PM"):
                
                if(len(datetime_list[1]) == 5):
                    datetime_list[1] = "0" + datetime_list[1]
                    hours = int(datetime_list[1][:2]) + 12
                    mins = datetime_list[1][2:4]
                elif(len(datetime_list[1]) == 6 and datetime_list[1][:2] == "12"):
                    hours = datetime_list[1][:2]
                    mins = datetime_list[1][2:4]
                else:
                    hours = int(datetime_list[1][:2]) + 12
                    mins = datetime_list[1][2:4]   
                
            elif(datetime_list[2] == "AM"):
                if(len(datetime_list[1]) == 5):
                    datetime_list[1] = "0" + datetime_list[1]
                    hours = datetime_list[1][:2]
                    mins = datetime_list[1][2:4]
                else:
                    hours = datetime_list[1][:2]
                    mins = datetime_list[1][2:4]
            
            date_time = year + month + date + str(hours) + str(mins)
            
        if col_name == "Status":
            status = value
            if status =="C/In":
                code = "0001"
            elif status =="C/Out":
                code = "0002"
                
    line = "31" + date_time + code + "00000" + ids + code

    with open("list.dat", "a") as my_file:
        my_file.write(line+"\n")
my_file.close


