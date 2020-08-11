import docx
from tkinter import *
import os
pwd=os.getcwd()
import sys
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
root=Tk()

root.title('Rent Invoice Automation')
# root.attributes('-fullscreen', True)
root.state('zoomed')
# sizex = 800
# sizey = 500
# posx  = 300
# posy  = 50
# root.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
try:
    root.wm_iconbitmap("C:\\Users\\darsh\\PycharmProjects\\test111\\test11\\New folder\\homelogo.ico")
    try:
        icon_path=resource_path('C:\\Users\\darsh\\PycharmProjects\\test111\\test11\\New folder\\homelogo.ico')
        root.iconbitmap(icon_path)
    except:
        pass
except:
    pass
# root.geometry('700x400')
# root.configure(bg='blue')
image_num=0
try:
    background_image = PhotoImage(file="C:\\Users\\darsh\\PycharmProjects\\test111\\test11\\New folder\\landscape.png")
    background_label = Label(root,image=background_image)
    background_label.place(x=0,y=0)
except:
    pass
# quitbutton=Button(root,text='X',bg='red',fg='white',activebackground='#00ff00',command=root.quit,bd=4,width=2,height=1)
# quitbutton.pack(side=TOP,anchor=NE)
# minimize_button=Button(root, text = "-",activebackground='#00ff00', command = lambda: root.wm_state("iconic"),width=2,height=1,bd=4).pack(side=TOP,anchor=NE)



def background():

    try:
        global background_label
        background_label.destroy()
        backgroundbutton.destroy()

    except:
        pass
# backgroundbutton = Button(root,text='Remove Background',activebackground='#00ff00',command=background)
# backgroundbutton.pack(side=BOTTOM)

def myfunction(event):
    canvas.configure(scrollregion=canvas.bbox("all"),width=700,height=400)

myframe = Frame(root,bd=5,width=700,height=600,relief=SUNKEN,highlightthickness=5,highlightcolor='black',highlightbackground='white')
myframe.place(x=300,y=0)

canvas = Canvas(myframe)
frame = Frame(canvas)

myscrollbar=Scrollbar(myframe,orient="vertical",command=canvas.yview)
canvas.configure(yscrollcommand=myscrollbar.set)
myscrollbar.pack(side="right",fill="y")
canvas.pack(side="right")
canvas.create_window((0,0),window=frame,anchor='nw')
frame.bind("<Configure>",myfunction)


datelabel=Label(frame,text='Date')
datelabel.grid(row=2,column=1)
date = Entry(frame,justify=CENTER,width=40,bd=4)
date.grid(row=2,column=2)

monthlabel=Label(frame,text='Invoice Month-Year')
monthlabel.grid(row=3,column=1)
month = Entry(frame,justify=CENTER,width=40,bd=4)
month.grid(row=3,column=2)

addresslabel=Label(frame,text='Address')
addresslabel.grid(row=4,column=1)
address = Entry(frame,justify=CENTER,width=40,bd=4)
address.grid(row=4,column=2)

rentlabel=Label(frame,text='Rent')
rentlabel.grid(row=5,column=1)
rent = Entry(frame,justify=CENTER,width=40,bd=4)
rent.grid(row=5,column=2)

petlabel=Label(frame,text='Pet fees')
petlabel.grid(row=6, column=1)
pet = Entry(frame,justify=CENTER,width=40,bd=4)
pet.grid(row=6,column=2)

watertrashlabel=Label(frame,text='Water and Trash Fees')
watertrashlabel.grid(row=7,column=1)
watertrash = Entry(frame,justify=CENTER,width=40,bd=4)
watertrash.grid(row=7,column=2)

seweragelabel=Label(frame,text='Sewerage fees')
seweragelabel.grid(row=8,column=1)
sewerage = Entry(frame,justify=CENTER,width=40,bd=4)
sewerage.grid(row=8,column=2)

electricitylabel=Label(frame,text='Electricity Charges')
electricitylabel.grid(row=9,column=1)
electricity = Entry(frame,justify=CENTER,width=40,bd=4)
electricity.grid(row=9,column=2)

otherslabel=Label(frame,text='Other Fees')
otherslabel.grid(row=10,column=1)
others = Entry(frame,justify=CENTER,width=40,bd=4)
others.grid(row=10,column=2)

input_savelabel=Label(frame,text='New file name without extension (Existing file might be overwritten)')
input_savelabel.grid(row=11,column=1)
input_save=Entry(frame,justify=CENTER,width=40,bd=4)
input_save.grid(row=11,column=2)

import sqlite3
conn = sqlite3.connect("test1.db")
cur = conn.cursor()

cur.execute('''
           CREATE TABLE IF NOT EXISTS invoices
           (date varchar(250), 
           invoice_month_year varchar(250),
           address varchar(250) ,
           rent decimal (50,4),
           pet_fees decimal (50,4),
           water_trash_fees decimal (50,4),
           sewerage_fees decimal (50,4),
           electricity_charges decimal (50,4),
           other_fees decimal (50,4),
           total decimal (50,4)
           )''')
# except:
#     pass

abcde=None
i=0
name_newlabel={}
amount_newlabel={}
name_new={}
amount_new = {}
a=12
def funct2():
    global i
    # if i <=5:

    global a
    i+=1
    a+=1
    name_newlabel[i] = Label(frame, text='Expense name')
    name_newlabel[i].grid(row=a,column=1)
    name_new[i] = Entry(frame,justify=CENTER,width=40,bd=4)
    name_new[i].grid(row=a,column=2)
    a+=1
    amount_newlabel[i] = Label(frame, text='Expense amount')
    amount_newlabel[i].grid(row=a,column=1)
    amount_new[i] = Entry(frame,justify=CENTER,width=40,bd=4)
    amount_new[i].grid(row=a,column=2)
        # if i==6:
        #     addbutton['state'] = 'disable'

def delfunction():
    global i
    try:
        name_newlabel[i].destroy()
        name_new[i].destroy()
        amount_newlabel[i].destroy()
        amount_new[i].destroy()
        i-=1
    except:
        pass
    try:
        addbutton['state'] = 'normal'
    except:
        pass
delbutton=Button(frame,text='Delete',activebackground='#00ff00',command=delfunction,justify=RIGHT,bd=4)
delbutton.grid(row=12,column=3)

addbutton=Button(frame,text='Add',activebackground='#00ff00',command=funct2,bd=4)
addbutton.grid(row=12,column=2)



def func1(date1,month1,address1,rent1,pet1,watertrash1,sewerage1,electricity1,others1,input_save1):

    global i
    global label123
    global file_exist_and_open

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
        folders = os.listdir()
        try:
            if "Invoices" not in folders:
                os.mkdir('Invoices')
            doc.save('Invoices\\'+doc_name)
        except OSError:
            doc.save(doc_name)
        date.destroy()
        labeldate=Label(frame,text=date1)
        labeldate.grid(row=2,column=2)
        month.destroy()
        labelmonth=Label(frame,text=month1)
        labelmonth.grid(row=3,column=2)
        address.destroy()
        labeladdress=Label(frame,text=address1)
        labeladdress.grid(row=4,column=2)
        rent.destroy()
        labelrent=Label(frame,text=rent1)
        labelrent.grid(row=5,column=2)
        pet.destroy()
        labelpet=Label(frame,text=pet1)
        labelpet.grid(row=6,column=2)
        watertrash.destroy()
        labelwatertrash=Label(frame,text=watertrash1)
        labelwatertrash.grid(row=7,column=2)
        sewerage.destroy()
        labelsewerage=Label(frame,text=sewerage1)
        labelsewerage.grid(row=8,column=2)
        electricity.destroy()
        labelelectricity=Label(frame,text=electricity1)
        labelelectricity.grid(row=9,column=2)
        others.destroy()
        labelothers=Label(frame,text=others1)
        labelothers.grid(row=10,column=2)
        input_save.destroy()
        global filename
        filename=input_save1
        labelinput_save=Label(frame,text=input_save1)
        labelinput_save.grid(row=11,column=2)
        addbutton.destroy()
        delbutton.destroy()
        global a
        a=12
        labelname_new={}
        labelamount_new={}

        for j in range(0,i):
            a+=1
            labelname_new[j]=Label(frame,text=abcdef[j][0])
            labelname_new[j].grid(row=a,column=1)
            labelamount_new[j]=Label(frame,text=abcdef[j][1])
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
        successlabel=Label(frame,text='Word file created in Invoices in same directory',fg='green')
        successlabel.grid(row=1,column=2)
        global uploadondrive
        uploadondrive = Button(frame, text='Upload on drive',activebackground='#00ff00', command=googledrive_entries)
        uploadondrive.grid(row=1,column=1)
        button1.destroy()

        #integrating with DB
        cur.execute("INSERT INTO invoices(date,invoice_month_year,address,rent,pet_fees,water_trash_fees,sewerage_fees,electricity_charges,other_fees,total) VALUES(?,?,?,?,?,?,?,?,?,?)",(date1,month1,address1,rent1,pet1,watertrash1,sewerage1,electricity1,others1,total))
        conn.commit()

        return doc_name
    except ValueError:
        try:
            file_exist_and_open.destroy()
        except:
            pass
        label123=Label(frame,text='Enter amount in integers or decimals only',fg='red')
        label123.grid(row=1,column=2)
        print('Characters allowed are [1,2,3,4,5,6,7,8,9,.]') #For debugging
    except OSError:
        try:
            label123.destroy()
        except:
            pass
        file_exist_and_open=Label(frame,text='Error:Word file is open. Close and submit',fg='red')
        file_exist_and_open.grid(row=1,column=2)





button1=Button(frame,text='Submit',bd=4,activebackground='#00ff00',command=lambda : func1(date.get(),month.get(),address.get(),rent.get(),pet.get(),watertrash.get(),sewerage.get(),electricity.get(),others.get(),input_save.get()))
button1.grid(row=1,column=1)



def googledrive_entries():
    global a
    folder_namelabel=Label(frame,text='Enter the folder name you want to create/present on drive: ')
    folder_namelabel.grid(row=a+2,column=1)
    global folder_name
    folder_name=Entry(frame,justify=CENTER,width=40,bd=4)
    folder_name.grid(row=a+2,column=2)
    file_namelabel=Label(frame,text='Enter the file name for the word document you want to give on drive: ')
    file_namelabel.grid(row=a+3,column=1)
    global title
    title=Entry(frame,justify=CENTER,width=40,bd=4)
    title.grid(row=a+3,column=2)

    fullpathlabel=Label(frame,text='Do not change path if file not moved ')
    fullpathlabel.grid(row=a+4,column=1)
    global fullpath
    fullpath=Entry(frame,justify=CENTER,width=40,bd=4)
    fullpath.grid(row=a+4,column=2)
    global filename,pwd
    fullpath.delete(0, END)
    pwd2slash=pwd.split('\\')
    path=''
    for j in pwd2slash[:-1]:
        path=path+j+'\\'
    finalpath=path+'Invoices\\'+filename+'.docx'
    fullpath.insert(0,finalpath)
    submitbutton.grid(row=a+5,column=1)
    a+=5
    global uploadondrive
    uploadondrive.destroy()


# GoogleAuth.DEFAULT_SETTINGS[
#     'client_config_file'] = 'C:\\Users\\darsh\\PycharmProjects\\test111\\client_secrets.json'


def check_folder_exists_and_create_fileupload(folder_name1,title1,fullpath1):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive
    GoogleAuth.DEFAULT_SETTINGS['client_config_file'] = "C:\\Users\\darsh\\PycharmProjects\\test111\\test11\\New folder\\client_secrets.json"

    # gauth.LoadCredentialsFile()
    gauth = GoogleAuth()
    # gauth.LocalWebserverAuth()
    # Creates local webserver and auto handles authentication.
    drive = GoogleDrive(gauth)
    global a
    def check_folder_exists(folder_name):
        list_of_file = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
        global folder_id
        list1=[]
        for drive_folder in list_of_file:
            list1.append(drive_folder['title'])
        if folder_name in list1:
            for drive_folder in list_of_file:
                if drive_folder['title'] == folder_name:

                    folder_id = drive_folder['id']
                    return folder_id

        else:
            folder = drive.CreateFile({'title': folder_name, "mimeType": "application/vnd.google-apps.folder"})
            folder.Upload()
            folder_id = folder['id']
            return folder_id
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
        successdrive=Label(frame,text='Successfully Uploaded on drive',fg='green')
        successdrive.grid(row=a,column=2)
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
        oserror=Label(frame,text='Path Error in uploading on drive',fg='red')
        oserror.grid(row=a,column=2)

submitbutton=Button(frame,text='Submit on drive',activebackground='#00ff00',command=lambda:check_folder_exists_and_create_fileupload(folder_name.get(),title.get(),fullpath.get()))

def sendmailfunction():
    sendmail.destroy()
    global fromadd,password,to
    global a
    # Warninglabel=Label(frame,text='If using gmail, Turn LESS SECURE APP access ON if it is OFF ')
    # Warninglabel.grid(row=a+1,column=1)
    # fromaddlabel=Label(frame, text='Email from')
    # fromadd=Entry(frame,justify=CENTER)
    # fromaddlabel.grid(row=a+2,column=1)
    # fromadd.grid(row=a+2,column=2)
    # passwordlabel=Label(frame,text='Password')
    # passwordlabel.grid(row=a+3,column=1)
    # password=Entry(frame,justify=CENTER)
    # password.grid(row=a+3,column=2)
    tolabel=Label(frame,text='Send to')
    tolabel.grid(row=a+2,column=1)
    to=Entry(frame,justify=CENTER,width=40,bd=4)
    to.grid(row=a+2,column=2)
    submitmail.grid(row=a+5, column=1)
    a+=2
def bindmailelements(from1,password1,to1):
    from email.message import EmailMessage

    fromowner = from1
    totenant = to1

    mssg = EmailMessage()
    mssg['From'] = fromowner
    mssg['To'] = totenant
    global monthyear
    mssg['Subject'] = 'Rent Invoice '+(monthyear)+ ' Uploaded on Drive'
    global folder_id
    body = ''' Hi, 

    Your monthly rent invoice has been uploaded on the drive. 
    https://drive.google.com/drive/folders/'''+str(folder_id)+ '''

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
        successmail=Label(frame,text='Mail sent',bg='green')
        successmail.grid(row=a+1,column=1)
    except Exception:
        try:
            successmail.destroy()
        except Exception:
            pass
        loginerror=Label(frame,text='Login Error in sending email',fg='red')
        loginerror.grid(row=a+1,column=1)
def bindmailelementsusingAPI(to1):
    from Google import Create_Service
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    global monthyear, a
    global folder_id
    CLIENT_SECRET_FILE = "C:\\Users\\darsh\\PycharmProjects\\test111\\test11\\New folder\\client_secrets.json"
    API_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    global authorizeerror,successsent,sendingerror

    try:
        service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

    except:
        try:
            successsent.destroy()
        except:
            pass
        try:
            sendingerror.destroy()
        except:
            pass
        authorizeerror = Label(frame, text='Authorization file moved', fg='red')
        authorizeerror.grid(row=a + 1, column=2)
    try:
        totenant = to1
        emailMsg = ''' Hi, 
    
        Your monthly rent invoice has been uploaded on the drive. 
        https://drive.google.com/drive/folders/'''+str(folder_id)+ '''
    
        Thanks. '''
        mimeMessage = MIMEMultipart()
        mimeMessage['to'] = totenant
        mimeMessage['subject'] = 'Rent Invoice '+(monthyear)+ ' Uploaded on Drive'
        mimeMessage.attach(MIMEText(emailMsg, 'plain'))
        raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

        message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
        print(message)
        try:
            authorizeerror.destroy()
        except:
            pass
        try:
            sendingerror.destroy()
        except:
            pass
        successsent = Label(frame, text='Mail successfully sent', fg='green')
        successsent.grid(row=a + 1, column=2)

    except:
        try:
            authorizeerror.destroy()
        except:
            pass
        try:
            successsent.destroy()
        except:
            pass
        sendingerror = Label(frame, text='Error in sending email', fg='red')
        sendingerror.grid(row=a + 1, column=2)

sendmail=Button(frame,text='Send Notification Email',activebackground='#00ff00',command=sendmailfunction)

# submitmail=Button(frame,text='Submit Email and Password',command=lambda:bindmailelements(fromadd.get(),password.get(),to.get()))#Use for SMTP

submitmail=Button(frame,text='Send email',activebackground='#00ff00',command=lambda:bindmailelementsusingAPI(to.get()))

root.mainloop()


