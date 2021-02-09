import paramiko
import pandas as pd
import getpass
from filecmp import dircmp
import os
import glob
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)


print("##################################################################################################################################")
print("NOTE:")
print("Please Create a folder COMPARE on your desktop and copy this script and PROD ICRF in the COMPARE folder and run this programme!!!!")
print("##################################################################################################################################")
created=input("Did you create the folder with name COMPARE (y/n) : ")
print("###################################################################################################################################")


if created=='y':

    ID=input("please enter your anthem ID : ")
    print('----------------------------------------------------------------------------------------------------------------------------------')

    #inputting the password
    pw = getpass.getpass()
    path_temp=os.getcwd()
    print('----------------------------------------------------------------------------------------------------------------------------------')

    #removing the files existing text file from the path
    path = f'{path_temp}'

    files_in_directory = os.listdir(path)
    filtered_files = [file for file in files_in_directory if file.endswith(".txt")]
    for file in filtered_files:
        path_to_file = os.path.join(path, file)
        os.remove(path_to_file)


    #clean up the UAT extracted folder

    if os.path.exists(f"{path}\\UAT"):
        try:
            files_in_directory = os.listdir(f'{path}\\UAT')
            filtered_files = [file for file in files_in_directory if file.endswith(".lst") or file.endswith(".sh") or file.endswith(".parm")]
            for file in filtered_files:
                path_to_file = os.path.join(f'{path}\\UAT', file)
                os.remove(path_to_file)
            os.rmdir(f"{path}\\UAT")
        except OSError:
            print (f"Removing of the directory {path}\\UAT failed")


    #clean up the PROD extracted folder

    if os.path.exists(f"{path}\\PROD"):
        try:
            files_in_directory = os.listdir(f'{path}\\PROD')
            filtered_files = [file for file in files_in_directory if file.endswith(".lst") or file.endswith(".sh") or file.endswith(".parm")]
            for file in filtered_files:
                path_to_file = os.path.join(f'{path}\\PROD', file)
                os.remove(path_to_file)
            os.rmdir(f"{path}\\PROD")
        except OSError:
            print (f"Removing of the directory {path}\\PROD failed")

    #finding the xls file in the currect path and passing as the input
    extension = 'xls'
    os.chdir(path)
    result_1 = glob.glob('*.{}'.format(extension))
    result=result_1[0]
    print("THE PRODUCTION ICRF: "+result)



    #using pamdas module and extracting the data of the files from the list 
    df = pd.ExcelFile(f"{path_temp}\\{result}").parse('ICRF_UNIX')
    names=[]
    UAT_path=[]
    PROD_path=[]
    names=df['Requestor US Domain ID'].tolist()
    names_list=names[4:]
    UAT_path=df['PlanView ID'].tolist()
    UAT_path_list=UAT_path[4:]

    SERVER=df['RITM'].tolist()
    SERVER_INFO=SERVER[4:]
    PROD_path=df['IM Imp Review Date'].tolist()
    PROD_path_list=PROD_path[4:]

    print('----------------------------------------------------------------------------------------------------------------------------------')

    # creating uat directories  
    if not os.path.exists(f"{path}\\UAT"):
        
        try:
            os.mkdir(f"{path}\\UAT")
        except OSError:
                print (f"Creation of the directory {path}\\UAT failed")
        else:
                print (f"Successfully created the directory {path}\\UAT ")
    
                print('----------------------------------------------------------------------------------------------------------------------------------')

    # creating uat directories 
    if not os.path.exists(f"{path}\\PROD"):
        
        try:
            os.mkdir(f"{path}\\PROD")
        except OSError:
                print (f"Creation of the directory {path}\\PROD failed")
        else:
                print (f"Successfully created the directory {path}\\PROD ")
                print('----------------------------------------------------------------------------------------------------------------------------------')

    UAT_host_comp='<server_1>'
    UAT_trg_path=f'{path_temp}\\UAT'

    PROD_host_comp='<server_2>'
    PROD_trg_path=f'{path_temp}\\PROD'

    DEV_host_comp='<server_3>'


    # #connecting to the servers

    for i,j,k in zip(names_list,UAT_path_list,SERVER_INFO):
        if k=='vaathmr381':
            ssh=paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(hostname=f'{UAT_host_comp}',username=ID,password=pw,port=22)
            sftp_client=ssh.open_sftp()
            try:
                sftp_client.get(f'{j}/{i}',f'{UAT_trg_path}\\{i}')
                f = open(f"{path_temp}\\uat_missing_files.txt", "a")
                f.close()
            except Exception  as e:
                print(str(e)+f" : file name : {i}")
                f = open(f"{path_temp}\\uat_missing_files.txt", "a")
                f.write(f"{i}")
                f.close()
            sftp_client.close()
            ssh.close()
        else:
            ssh=paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(hostname=f'{DEV_host_comp}',username=ID,password=pw,port=22)
            sftp_client=ssh.open_sftp()
            try:
                sftp_client.get(f'{j}/{i}',f'{UAT_trg_path}\\{i}')
                f = open(f"{path_temp}\\uat_missing_files.txt", "a")
                f.close()
            except Exception  as e:
                print(str(e)+f" : file name : {i}")
                f = open(f"{path_temp}\\uat_missing_files.txt", "a")
                f.write(f"{i}")
                f.write("\n")
                f.close()
            sftp_client.close()
            ssh.close()
        
    print("########################################################################################################################")

    ssh=paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname=f'{PROD_host_comp}',username=ID,password=pw,port=22)
    sftp_client=ssh.open_sftp()
    
    for i,j in zip(names_list,PROD_path_list):
        try:
            sftp_client.get(f'{j}/{i}',f'{PROD_trg_path}\\{i}')
            f = open(f"{path_temp}\\prod_missing_files.txt", "a")
            f.close()
        except Exception  as e:
            print(str(e)+f" : file name : {i}")
            f = open(f"{path_temp}\\prod_missing_files.txt", "a")
            f.write(f"{i}")
            f.write("\n")
            f.close()
    
    sftp_client.close()
    ssh.close()

    print("########################################################################################################################")


    #check for the empty files in the directory

    def delete_list(path):
        files_tgt_list=os.listdir(path)
        for i in files_tgt_list:
            if os.stat(f"{path}\\{i}").st_size == 0:
                print(f'{i} : file is empty')
                os.remove(f"{path}\\{i}")

    #delete the empty files in the uat path
    print("removing UAT files if empty!!!")
    delete_list(UAT_trg_path)
    print("########################################################################################################################")
    # #deleting all the empty files in the prod loc
    print("removing PROD files if empty!!!")
    delete_list(PROD_trg_path)


    print("########################################################################################################################")


    #function used to find the difference between two folders
    def print_diff_files(dcmp):
        for name in dcmp.diff_files:
            return "difference in  %s found between %s and %s" % (name, dcmp.left,dcmp.right)
    dcmp = dircmp(UAT_trg_path,PROD_trg_path) 

    info = print_diff_files(dcmp)


    if info == None:
        info = "No difference between two folders"


    print("########################################################################################################################")


    def missing_files(uat,prod):
        list_uat = os.listdir(uat)
        list_prod = os.listdir(prod)
        result_list=[]
        for i in list_uat + list_prod:
            if i not in list_uat or i not in list_prod:
                result_list.append(i)
        if len(result_list) == 0:
            print("NO DIFFERENCE IN DIRECTORIES")
            message="NO DIFFERENCE IN DIRECTORIES!!!"
        else:
            print("THESE FILES ARE MISSING IN PROD DIRECTORY: "+ str(result_list))
            message=f"THESE FILES ARE MISSING IN PROD DIRECTORY: {result_list}"
        return message

    print("##################################################################################################################################")

    #extract mailing details from outlook
    outlookMail = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlookMail.Folders.Item(1)
    selfName=root_folder.Name

    
    #this field is optional
    # To attach a file to the email (optional):

    if os.path.exists(f'{path}\\prod_missing_files.txt') and os.stat(f"{path}\\prod_missing_files.txt").st_size != 0 :
        attachment  = f'{path_temp}\\prod_missing_files.txt'
        #emailing the details 
        mail.To = f'{selfName}'
        mail.Subject = 'THE DIFFERENCE IN ICRF PRODUCTION VALIDATION'
        mail.Body = 'Message body'
        mail.HTMLBody = f'''<h2>THE FILE WITH DIFFERENCE ARE BELOW:- </h2>
        <h2>------------------------------------------------------------------------</h2>
        <p1> {info} </p1><h2>THE MISSING FILES ARE BELOW:- </h2>
        <h2>------------------------------------------------------------------------</h2>
        <p1> {missing_files(UAT_trg_path,PROD_trg_path)} </p1>
        <h2>------------------------------------------------------------------------</h2>'''
        mail.Attachments.Add(attachment)
        mail.Send()
    elif os.path.exists(f'{path}\\uat_missing_files.txt') and os.stat(f"{path}\\uat_missing_files.txt").st_size != 0:
        attachment  = f'{path_temp}\\uat_missing_files.txt'
                    #emailing the details 
        mail.To = f'{selfName}'
        mail.Subject = 'THE DIFFERENCE IN ICRF PRODUCTION VALIDATION'
        mail.Body = 'Message body'
        mail.HTMLBody = f'''<h2>THE FILE WITH DIFFERENCE ARE BELOW:- </h2>
        <h2>------------------------------------------------------------------------</h2>
        <p1> {info} </p1><h2>THE MISSING FILES ARE BELOW:- </h2>
        <h2>------------------------------------------------------------------------</h2>
        <p1> {missing_files(UAT_trg_path,PROD_trg_path)} </p1>
        <h2>------------------------------------------------------------------------</h2>'''
        mail.Attachments.Add(attachment)
        mail.Send()
    elif os.stat(f"{path}\\uat_missing_files.txt").st_size == 0 and os.stat(f"{path}\\prod_missing_files.txt").st_size == 0:
                    #emailing the details 
        mail.To = f'{selfName}'
        mail.Subject = 'THERE IS NO DIFFERENCE IN ICRF PRODUCTION VALIDATION'
        mail.Body = 'Message body'
        mail.HTMLBody = f'''<h2>NO DIFFERENCE DETECTED </h2>'''
        mail.Send()
        os.remove(f"{path}\\uat_missing_files.txt")
        os.remove(f"{path}\\prod_missing_files.txt")
    
    else:
        pass


    #removing the text files
    files_in_directory = os.listdir(path)
    filtered_files = [file for file in files_in_directory if file.endswith(".txt")]
    for file in filtered_files:
        path_to_file = os.path.join(path, file)
        os.remove(path_to_file)



else:
    print("##################################################################################################################################")
    print("NOTE:")
    print("Please Create a folder COMPARE on your desktop and copy this script and PROD ICRF in the COMPARE folder and run this programme!!!!")
    print("##################################################################################################################################")
    pass