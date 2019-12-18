import urllib
import requests
import datetime
import xlsxwriter
import ctypes
import json
import pyodbc
import textwrap
import time





    
def getStartEntry():
    fopen = open("PythonUtil/lastEntry.txt", 'r')
    entry = int(fopen.read())
    fopen.close()
    return entry

def getFileName():
    now = datetime.datetime.now()
    return now.strftime('%y%m%d %I-%M-%S.xlsx')

def writeStartEntry():
    maxEntry = getMaxEntry()
    fopen = open("PythonUtil/lastEntry.txt", 'w')    
    fopen.write(str(maxEntry - 100))
    fopen.close()

def getMaxEntry():
    global data

    entryID = list()
    for point in data:
        entryID.append(int(point.get('EntryId')))
    

    return max(entryID)

def getHeaderList():
    fopen = open("PythonUtil/headers.txt", 'r')
    lines = fopen.read().splitlines()
    return lines

def getValueString(numValues):
    values = '?'
    for i in range(1, numValues):
        values += ", ?"
    return values

def getColumnList():
    fopen = open("PythonUtil/values.txt", 'r')
    columns = fopen.read().splitlines()
    

    return columns
    
def writeToXlsx(headers):
    global entries

    
    col = 0
    for header in headers:
        worksheet.write(0, col, header)
        col += 1


    row = 0
    col = 0


    ##for key in data[0].keys():
        ##worksheet.write(row, col, key)
        ##col += 1
    for entry in entries:
        row += 1
        col = 0
        for key in entry:
            worksheet.write(row, col, entry.get(key))
            col += 1

    workbook.close()

def importToAccess(filename, columns, values):
    # Create connection to Special Access Database
    DBconn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=W:\ops\BAF\Technology\Databases\SpecialAccessDB\Special Access Database.accdb;')
    DBcursor = DBconn.cursor()

    # Create connection to newly-created Excel file
    exFile = 'W:\ops\BAF\Technology\Databases\SpecialAccessDB\Wufoo Data Files\ '.strip() + filename
    exconn = pyodbc.connect(r'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ='+exFile,autocommit=True)
    excursor = exconn.cursor()

    excelResults = excursor.execute('select * from [Sheet1$]').fetchall()

    #Create SQL string for data insertion
    sql = 'insert into [Wufoo Form Data]('      # Start of string
    for c in columns:                           # Concatenate all column names in string
        if c == columns[len(columns)-1]:
            sql += c
        else:
            sql += c + ", "

    sql += ') values(' + values + ')'        # Insert values string of ?s account for insertion (See pyodbc documentation for details)
    errors = 0
    for row in excelResults:                    # Insert data row by row
        try:
            DBcursor.execute(sql, row)
        except:
            errors += 1                         # If key already exists, add one to error
    DBconn.commit()                             # Commit all SQL statements in DBcursor  
    return errors                               # Return number of errors for suer feedback
   

def messageBox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


if __name__ == '__main__':

    print("SPECIAL ACCESS Database Automated Data Collection")
    print("-----------------------------------------------")


    # 'Global' variables
    WUFOO_FORM = 'forms/qzitjkc1yixswz/entries.json?pageStart=%s&pageSize=100' # Form information (which form to pull from, how many entries to pull)
    NUM_VALUES = 38 # Number of columns in table
    values = getValueString(NUM_VALUES) # Get string of '?' for SQL insertion
    headerList = getHeaderList() # Get list of headers for Excel file
    columns = getColumnList() # Get string of columns for SQL insertion


    # Information to authenticate Wufoo login
    base_url = '-------------' #Redacted
    username = '-------------' #Redacted
    password = '-------------' #Redacted

    # Authenticate Wufoo login
    print("Authenticating Wufoo access...")
    password_manager = urllib.request.HTTPPasswordMgrWithDefaultRealm()
    password_manager.add_password(None, base_url, username, password)
    handler = urllib.request.HTTPBasicAuthHandler(password_manager)
    opener = urllib.request.build_opener(handler)

    urllib.request.install_opener(opener)

    # Retrieve last 100 entries from given form
    print('Retrieving Wufoo data...')    
    numEntries = getStartEntry()
    form_string = WUFOO_FORM % str(numEntries) #Get entries from form starting from entry "numEntries"
    response = urllib.request.urlopen(base_url+form_string)
    result = json.load(response)
    data = result['Entries'] #List of entries retrieved from Wufoo form
    entries = list()
    for entry in data:
        if entry["DateCreated"] < "2019-07-15":
            pass
        else:
            entries.append(entry)
            
     
    # Update most recent entry number in text file for next program execution
    fileName = getFileName()
    print('Updating most recent entry number...')
    writeStartEntry()

    
    #Create XLSX workbook and write data to Excel
    workbook  = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()
    print('Writing data to Excel...')
    writeToXlsx(headerList)
    
    #Import data from Excel to Access,        
    print("Importing data to Access database...")
    errors = importToAccess(fileName, columns, values)

    

    messageBox('Special Access Database', 'Done! ' + str(len(entries) - errors) + ' new records added.' , 1)
            

    
    


            
            
