import os
import glob
import pandas as pd
import teradata
import numpy as np
import warnings
import win32com.client as win32
import getpass
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
warnings.filterwarnings("ignore")

#removing any previous log files and directory
path=os.getcwd()
if os.path.exists(f"{path}\\logs"):
    try:
        files_in_directory = os.listdir(f'{path}\\logs')
        filtered_files = [file for file in files_in_directory if file.endswith(".log")]
        for file in filtered_files:
            path_to_file = os.path.join(f'{path}\\logs', file)
            os.remove(path_to_file)
        os.rmdir(f"{path}\\logs")
    except OSError:
        print (f"Removing of the directory {path}\\logs failed")

#defining all the parameters used for process
path=os.getcwd()
extension = 'xlsm'
os.chdir(path)
result_1 = glob.glob('*.{}'.format(extension))
result = result_1[0]
print("#########################################################################")
print("THE DB WORKBOOK IS : " + str(result))
print("#########################################################################")
ID=input("please enter your anthem ID : ")
print("#########################################################################")
pw = getpass.getpass()
print("#########################################################################")
df = pd.ExcelFile(f"{path}\\{result}").parse('DB Objects')
table_names=df['DVLPR.2'].tolist()
table_names_list=table_names[3:]



res = []
for i in table_names_list:
    if i not in res:
        res.append(i)

# Establish the connection to the Teradata database
udaExec = teradata.UdaExec(appName="test", version="1.0", logConsole=False)
connection = udaExec.connect(method="odbc", system="<servername>", authentication='LDAP',
                                 driver="Teradata Database ODBC Driver 16.20",
                                 username=f'{ID}', password=f'{pw}');

#writing to a file 
for Table_name in res[:-1]:
    query =  f'''SELECT ColumnName,ColumnType                    
    FROM DBC.Columns
    WHERE DatabaseName='MRDM'
    AND TableName='{Table_name}'
    order by ColumnName'''
    f = open("IR_ENV.txt", "a+")
    for row in connection.execute(query):
        f.write(str(row))
        f.write('\n')
f.close()

#closes teradata connection
connection.close()

# Establish the connection to the Teradata database
udaExec = teradata.UdaExec(appName="test", version="1.0", logConsole=False)
connection = udaExec.connect(method="odbc", system="<servername>", authentication='LDAP',
                             driver="Teradata Database ODBC Driver 16.20",
                             username=f'{ID}', password=f'{pw}');


#writing to a file 
for Table_name in res[:-1]:
    query = f'''SELECT ColumnName,ColumnType   
    FROM DBC.Columns
    WHERE DatabaseName='<DBNAME>'
    AND TableName='{Table_name}'
    order by ColumnName'''

    f = open("PROD_ENV.txt", "a+")
    for row in connection.execute(query):
        f.write(str(row))   
        f.write('\n')
f.close()


#closes teradata connection
connection.close()


#detect any difference if present 
with open(f'{path}\\IR_ENV.txt', 'r') as file1:
    with open(f'{path}\\PROD_ENV.txt', 'r') as file2:
        difference = set(file1).difference(file2)

difference.discard('\n')

with open('diff.txt', 'w') as file_out:
    for line in difference:
        file_out.write(line)



#extract mailing details from outlook
outlookMail = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlookMail.Folders.Item(1)
selfName=root_folder.Name

        

if os.stat(f"{path}\\diff.txt").st_size == 0:
    info="DB WORKBOOK VALIDATED SUCCESSFULLY"
    attachment  = f'{path}\\IR_ENV.txt'
    attachment_1  = f'{path}\\PROD_ENV.txt'
                    #emailing the details 
    mail.To = f'{selfName}'
    mail.Subject = 'NO DIFFERENCE DETECTED IN DB WORK BOOK'
    mail.Body = 'Message body'
    mail.HTMLBody = f'''<h2>------------------------------------------------------------------------</h2>
    <p1> {info} </p1>
    <h2>------------------------------------------------------------------------</h2>'''
    mail.Attachments.Add(attachment)
    mail.Attachments.Add(attachment_1)
    mail.Send()
else:
    info="DB WORKBOOK VALIDATED HAS MISMATCH COLUMNS OR MIGRATION IS NOT COMPLETED"
    attachment  = f'{path}\\IR_ENV.txt'
    attachment_1  = f'{path}\\PROD_ENV.txt'
                #emailing the details 
    mail.To = f'{selfName}'
    mail.Subject = 'THE DIFFERENCE IN DB WORK BOOK DETECTED'
    mail.Body = 'Message body'
    mail.HTMLBody = f'''<h2>ATTCHING FILES WITH DIFFERENCE DIFFERENCE ARE BELOW:- </h2>
    <h2>------------------------------------------------------------------------</h2>
    <p1> {info} </p1>
    <h2>------------------------------------------------------------------------</h2>'''
    mail.Attachments.Add(attachment)
    mail.Attachments.Add(attachment_1)
    mail.Send()


#removing the text fies
files_in_directory = os.listdir(path)
filtered_files = [file for file in files_in_directory if file.endswith(".txt") or file.endswith(".runNumber")]
for file in filtered_files:
    path_to_file = os.path.join(path, file)
    os.remove(path_to_file)





#removing the logs directory
if os.path.exists(f"{path}\\logs"):
    try:
        files_in_directory = os.listdir(f'{path}\\logs')
        filtered_files = [file for file in files_in_directory if file.endswith(".log")]
        for file in filtered_files:
            path_to_file = os.path.join(f'{path}\\logs', file)
            os.remove(path_to_file)
        os.rmdir(f"{path}\\logs")
    except OSError:
        # print (f"Removing of the directory {path}\\logs failed")
        print (f"check the directory {path}\\logs for more details")
        


