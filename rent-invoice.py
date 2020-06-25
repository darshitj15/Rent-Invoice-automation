import docx
import googleapiclient
from tkinter import *
import os
pwd=os.getcwd()

root=Tk()
root.title('Rent Invoice Automation')

datelabel=Label(root,text='Date')
datelabel.grid(row=2,column=1)
date = Entry(root)
date.grid(row=2,column=2)

monthlabel=Label(root,text='Invoice Month-Year')
monthlabel.grid(row=3,column=1)
month = Entry(root)
month.grid(row=3,column=2)

addresslabel=Label(root,text='Address')
addresslabel.grid(row=4,column=1)
address = Entry(root)
address.grid(row=4,column=2)

rentlabel=Label(root,text='Rent')
rentlabel.grid(row=5,column=1)
rent = Entry(root)
rent.grid(row=5,column=2)

petlabel=Label(root,text='Pet fees')
petlabel.grid(row=6, column=1)
pet = Entry(root)
pet.grid(row=6,column=2)

watertrashlabel=Label(root,text='Water and Trash Fees')
watertrashlabel.grid(row=7,column=1)
watertrash = Entry(root)
watertrash.grid(row=7,column=2)

seweragelabel=Label(root,text='Sewerage fees')
seweragelabel.grid(row=8,column=1)
sewerage = Entry(root)
sewerage.grid(row=8,column=2)

electricitylabel=Label(root,text='Electricity Charges')
electricitylabel.grid(row=9,column=1)
electricity = Entry(root)
electricity.grid(row=9,column=2)

otherslabel=Label(root,text='Other Fees')
otherslabel.grid(row=10,column=1)
others = Entry(root)
others.grid(row=10,column=2)

input_savelabel=Label(root,text='New file name without extension (Existing file might be overwritten)')
input_savelabel.grid(row=11,column=1)
input_save=Entry(root)
input_save.grid(row=11,column=2)


abcde=None
i=0
name_newlabel={}
amount_newlabel={}
name_new={}
amount_new = {}
a=12
def funct2():
    global i
    global a
    i+=1
    a+=1
    name_newlabel[i] = Label(root, text='Expense name')
    name_newlabel[i].grid(row=a,column=1)
    name_new[i] = Entry(root)
    name_new[i].grid(row=a,column=2)
    a+=1
    amount_newlabel[i] = Label(root, text='Expense amount')
    amount_newlabel[i].grid(row=a,column=1)
    amount_new[i] = Entry(root)
    amount_new[i].grid(row=a,column=2)
def delfunction():
    global i
    name_newlabel[i].destroy()
    name_new[i].destroy()
    amount_newlabel[i].destroy()
    amount_new[i].destroy()
    i-=1
delbutton=Button(root,text='Delete',command=delfunction)
delbutton.grid(row=12,column=3)

addbutton=Button(root,text='Add',command=funct2)
addbutton.grid(row=12,column=2)



def func1(date1,month1,address1,rent1,pet1,watertrash1,sewerage1,electricity1,others1,input_save1):

    global i
    global label123
    global file_exist_and_open
    addbutton.destroy()
    delbutton.destroy()
    abcdef=[]

    try:
        for j in range(1,i+1):
            abc = name_new[j].get()
            abcd = amount_new[j].get()
            abcde=[abc,abcd]
            abcdef.append(abcde)
    except:
        pass
    global monthyear
    monthyear=month1
    doc = docx.Document()
    doc.add_heading('Date: ' + date1, 5)
    doc.add_heading('Invoice Month:' + month1, 5)
    doc.add_heading('Property Address: ' + address1, 5)
    doc.add_paragraph()
    doc.add_paragraph()
    records = [['rent', rent1],
               ['Pet fees', pet1],
               ['Water & Trash (City of Sacramento)', watertrash1],
               ['Sewerage (Sac County)', sewerage1],
               ['Electricity (SMUD)', electricity1],
               [' Others (if any)', others1]]
    if abcdef!=[]:
        for j in abcdef:
            records.append(j)
    try:
        expense_values=[float(i[1]) for i in records]
        total=sum(expense_values)
        table1 = doc.add_table(rows=0, cols=2)
        table1.style = 'Table Grid'
        for reason, amount in records:
            row_Cells = table1.add_row().cells
            row_Cells[0].text = reason
            row_Cells[1].text = str(amount)
        row_Cells = table1.add_row().cells
        row_Cells[0].text = 'Total'
        row_Cells[1].text = str(total)
        doc_name = input_save1 + '.docx'
        doc.save(doc_name)

        date.destroy()
        labeldate=Label(root,text=date1)
        labeldate.grid(row=2,column=2)
        month.destroy()
        labelmonth=Label(root,text=month1)
        labelmonth.grid(row=3,column=2)
        address.destroy()
        labeladdress=Label(root,text=address1)
        labeladdress.grid(row=4,column=2)
        rent.destroy()
        labelrent=Label(root,text=rent1)
        labelrent.grid(row=5,column=2)
        pet.destroy()
        labelpet=Label(root,text=pet1)
        labelpet.grid(row=6,column=2)
        watertrash.destroy()
        labelwatertrash=Label(root,text=watertrash1)
        labelwatertrash.grid(row=7,column=2)
        sewerage.destroy()
        labelsewerage=Label(root,text=sewerage1)
        labelsewerage.grid(row=8,column=2)
        electricity.destroy()
        labelelectricity=Label(root,text=electricity1)
        labelelectricity.grid(row=9,column=2)
        others.destroy()
        labelothers=Label(root,text=others1)
        labelothers.grid(row=10,column=2)
        input_save.destroy()
        global filename
        filename=input_save1
        labelinput_save=Label(root,text=input_save1)
        labelinput_save.grid(row=11,column=2)
        global a
        a=12
        labelname_new={}
        labelamount_new={}

        for j in range(0,i):
            a+=1
            labelname_new[j]=Label(root,text=abcdef[j][0])
            labelname_new[j].grid(row=a,column=1)
            labelamount_new[j]=Label(root,text=abcdef[j][1])
            labelamount_new[j].grid(row=a,column=2)
            name_new[j+1].destroy()
            amount_new[j+1].destroy()
            name_newlabel[j+1].destroy()
            amount_newlabel[j+1].destroy()
        try:
            label123.destroy()
        except Exception:
            pass
        try:
            file_exist_and_open.destroy()
        except Exception:
            pass
        successlabel=Label(root,text='File successfully created on PC',fg='green')
        successlabel.grid(row=1,column=2)
        global uploadondrive
        uploadondrive = Button(root, text='Upload on drive', command=googledrive_entries)
        uploadondrive.grid(row=1,column=1)
        button1.destroy()

        return doc_name
    except ValueError:
        try:
            file_exist_and_open.destroy()
        except:
            pass
        label123=Label(root,text='Enter amount in integers or decimals only',fg='red')
        label123.grid(row=1,column=2)
        print('Characters allowed are [1,2,3,4,5,6,7,8,9,.]') #For debugging
    except OSError:
        try:
            label123.destroy()
        except:
            pass
        file_exist_and_open=Label(root,text='The word file you mentioned is open. Close to update ',fg='red')
        file_exist_and_open.grid(row=1,column=2)





button1=Button(root,text='Submit',command=lambda : func1(date.get(),month.get(),address.get(),rent.get(),pet.get(),watertrash.get(),sewerage.get(),electricity.get(),others.get(),input_save.get()))
button1.grid(row=1,column=1)

quitbutton=Button(root,text='Exit',command=root.quit)
quitbutton.grid(row=0,column=3)


def googledrive_entries():
    global a
    folder_namelabel=Label(root,text='Enter the folder name you want to create/present on drive: ')
    folder_namelabel.grid(row=a+2,column=1)
    global folder_name
    folder_name=Entry(root)
    folder_name.grid(row=a+2,column=2)
    file_namelabel=Label(root,text='Enter the file name for the word document you want to give on drive: ')
    file_namelabel.grid(row=a+3,column=1)
    global title
    title=Entry(root)
    title.grid(row=a+3,column=2)
    fullpathlabel=Label(root,text='Change path of the file if moved: ')
    fullpathlabel.grid(row=a+4,column=1)
    global fullpath
    fullpath=Entry(root)
    fullpath.grid(row=a+4,column=2)
    global filename,pwd
    fullpath.delete(0, END)
    pwd2slash=pwd.split('\\')
    path=''
    for j in pwd2slash:
        path=path+j+'\\'
    finalpath=path+filename+'.docx'
    fullpath.insert(0,finalpath)
    submitbutton.grid(row=a+5,column=1)
    a+=5
    global uploadondrive
    uploadondrive.destroy()



def check_folder_exists_and_create_fileupload(folder_name1,title1,fullpath1):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    # GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = 'C:\\Users\\darsh\\PycharmProjects\\test111\\client_secrets.json'
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()  # Creates local webserver and auto handles authentication.
    drive = GoogleDrive(gauth)

    def check_folder_exists(folder_name):
        list_of_file = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()

        list1=[]
        for drive_folder in list_of_file:
            list1.append(drive_folder['title'])
        if folder_name in list1:
            for drive_folder in list_of_file:
                if drive_folder['title'] == folder_name:
                    return drive_folder['id']

        else:
            folder = drive.CreateFile({'title': folder_name, "mimeType": "application/vnd.google-apps.folder"})
            folder.Upload()
            return folder['id']
    fid1 = check_folder_exists(folder_name1)

    def upload_file_inside_folder(title, fid1, fullpath):
        file = drive.CreateFile({'title': title, 'parents': [{'kind': 'drive#fileLink', 'id': fid1}]})
        file.SetContentFile(fullpath)
        file.Upload()
        return True
    global oserror
    global successdrive
    try:
        upload_file_inside_folder(title1, fid1, fullpath1)
        successdrive=Label(root,text='Successfully Uploaded on drive',fg='green')
        successdrive.grid(row=1,column=1)
        sendmail.grid(column=1)
        try:
            oserror.destroy()
        except:
            pass
    except Exception:
        try:
            successdrive.destroy()
        except Exception:
            pass
        oserror=Label(root,text='Path Error in uploading on drive',fg='red')
        oserror.grid(row=1,column=1)


submitbutton=Button(root,text='Submit on drive',command=lambda:check_folder_exists_and_create_fileupload(folder_name.get(),title.get(),fullpath.get()))
def sendmailfunction():
    sendmail.destroy()
    global fromadd,password,to
    global a
    Warninglabel=Label(root,text='If using gmail, Turn LESS SECURE APP access ON if it is OFF ')
    Warninglabel.grid(row=a+1,column=1)
    fromaddlabel=Label(root, text='Email from')
    fromadd=Entry(root)
    fromaddlabel.grid(row=a+2,column=1)
    fromadd.grid(row=a+2,column=2)
    passwordlabel=Label(root,text='Password')
    passwordlabel.grid(row=a+3,column=1)
    password=Entry(root)
    password.grid(row=a+3,column=2)
    tolabel=Label(root,text='Send to')
    tolabel.grid(row=a+4,column=1)
    to=Entry(root)
    to.grid(row=a+4,column=2)
    submitmail.grid(row=a+5, column=2)
    a+=5
def bindmailelements(from1,password1,to1):
    from email.message import EmailMessage

    fromowner = from1
    totenant = to1

    mssg = EmailMessage()
    mssg['From'] = fromowner
    mssg['To'] = totenant
    global monthyear
    mssg['Subject'] = 'Rent Invoice '+(monthyear)+ ' Uploaded on Drive'

    body = ''' Hi, 

    Your monthly rent invoice has been uploaded on the drive. 

    Thanks. '''

    mssg.set_content(body)

    import smtplib
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    global successmail,loginerror
    global a
    try:
        server.login(fromowner, password1)
        text = mssg.as_string()
        server.sendmail(fromowner, totenant, text)
        server.quit()
        try:
            loginerror.destroy()
        except Exception:
            pass
        successmail=Label(root,text='Mail sent',bg='green')
        successmail.grid(row=a+1,column=1)
    except Exception:
        try:
            successmail.destroy()
        except Exception:
            pass
        loginerror=Label(root,text='Login Error in sending email',bg='red')
        loginerror.grid(row=a+1,column=1)

sendmail=Button(root,text='Send Notification Email',command=sendmailfunction)
submitmail=Button(root,text='Submit Email and Password',command=lambda:bindmailelements(fromadd.get(),password.get(),to.get()))
root.mainloop()


