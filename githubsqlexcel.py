#Not the cleanest code, but I'll optimize it to work with a GUI

import pyodbc
import pandas as pd
import openpyxl

#from datatime import datetime 
'exec(%matplotlib inline)'


names = ('')
serverNames = ('')
db = ('')
#list
print("Choose server and database from the following list." )
print('Types: ') 
print(names)
print('Server: ') 
print(serverNames)
print('Database: ') 
print(db)
#choices 
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
serverName = input("Enter server name: ")  #input

databaseName = input("Enter database name: ") #input

query = pyodbc.connect('Driver={SQL Server};' 'Server=' + serverName + ';' + 'Database=' + databaseName + ';' + 'Trusted_Connection=yes;')
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
#Will ask in console to paste your SQL query 
sql = input("Input SQL Query(Check server & database): ") #input
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
#Read the SQL query for the driver, server, and database connection
data = pd.read_sql(sql, query)

#Formats table into dataframe using the read_sql function 
Data = pd.DataFrame(data)
print(Data)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------
#Exports database table to Excel format 
file = input("Create file name: ") #input 
newFile = ("C:\\Users\\...insert path" + file + ".xlsx")
export_excel = Data.to_excel(newFile, index=None, header=True) 
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------
wb = openpyxl.load_workbook(filename = newFile) #uses the openpyxl module by loading the workbook 
worksheet = wb.active #activates the worksheet -- need to make sure to activate wb.save() at the end if changing excel 

for col in worksheet.columns:
    max_length = 0
    column = col[0].column

    for cell in col:
        try: 
#based on the value of the cell, it makes sure it equals the max_length 
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

wb.save(newFile)

print("File has been created!")
#----------------------------------------------------------------------------------------------------------------------------------------------------------



