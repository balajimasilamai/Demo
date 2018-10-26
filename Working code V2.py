import tkinter
from tkinter import *
import cx_Oracle
from tkinter import messagebox
import csv 
from functools import partial
#======================================================================
val='Run'
root=Tk()
root.title('Automation Tool')
listofmodules=[]
module_name=[]
listofbackbuttons=[]
count=0
dict1={'Test Case Description':'Result'}
filename=''
#==============================================
def connect():
    global connection
    try:
        #messagebox.showinfo('Connection status','Connecting....')
        connection = cx_Oracle.connect(Entry1.get(),Entry2.get(),Entry3.get()+'/'+Entry4.get())
        messagebox.showinfo('Connection status','Connection established successfully')
        for i in range(0,len(list_buttons),1):
            list_buttons[i].config(state=NORMAL)
        
    except cx_Oracle.DatabaseError as e:
        error ,=e.args
        messagebox.showerror('Error',error   )
#############################################################
def user_validation():
    if name_entry.get().lower() == 'balaji' and pwd_entry.get()=='BALAJI' :
        #messagebox.showinfo('Logged in','Click Next to continue')
        #Next.config(state=NORMAL)
        login_frame.grid_forget()
        main_frame.tkraise()
        main_frame.grid(row=0,column=0,sticky=W+E+N+S)
    else:
        messagebox.showerror('Waring','Inavalid Credentials'   )
        #Next.config(state=DISABLED)
#============================
def quit_program():
    root.destroy()
def back_to_login_frame():
    login_frame.grid(row=0,column=0,sticky=W+E+N+S)
    main_frame.grid_forget()
##############################LOGIN FRAME############################################
login_frame=Frame(root,height=30,width=500,borderwidth=10,bg='Powder Blue')
login_frame.grid(row=0,column=0,sticky=W+N+E+S)
log_text=Label(login_frame,text='Login Info',font=('Arial',10,'bold'))
log_text.grid(row=11,column=0,sticky=W)
empty=Label(login_frame,text='                                                                  ',bg='Powder Blue')
empty.grid(row=11,columnspan=50,sticky=E)
name=Label(login_frame,text='User Name')
name.grid(row=12,column=0,sticky=W,padx=6)
name_entry=Entry(login_frame,bd=5)
name_entry.grid(row=12,column=1,sticky=W,padx=6)
pwd=Label(login_frame,text='Password')
pwd.grid(row=13,column=0,sticky=W,padx=6)
pwd_entry=Entry(login_frame,bd=5,show='*')
pwd_entry.grid(row=13,column=1,sticky=W,padx=6)
submit=Button(login_frame,text='Login',command=user_validation)

submit.grid(row=14,column=1,sticky=W,padx=6,pady=4,ipadx=4)

Quit=Button(login_frame,text='Quit',command=quit_program)

Quit.grid(row=14,column=2,sticky=W,padx=6,pady=4,ipadx=4)
################################MAIN FRAME#################################################
main_frame=Frame(root,height=30,width=500,borderwidth=10,bg='Powder Blue')
main_frame.grid(row=0,column=0,sticky=W+E+N+S)
back=Button(root,text='Back')
back.grid(row=40,column=0,sticky=W)
Db_details=Frame(main_frame,height=20,width=50,highlightthickness=3)
Db_details.grid(row=0,column=0,sticky=W)
label=Label(Db_details,text='Database information',font=('Arial',10,'bold'),fg='blue')
label.grid(row=1,column=0,sticky=W)

User_name=Label(Db_details,text='User name',padx=3,font=('Arial',10,'bold'))
User_name.grid(row=2,column=0)
Entry1=Entry(Db_details,bd=5)
Entry1.grid(row=2,column=1)
Password=Label(Db_details,text='Password',padx=3,font=('Arial',10,'bold'))
Password.grid(row=3,column=0)
Entry2=Entry(Db_details,bd=5,show='*')
Entry2.grid(row=3,column=1)
Host=Label(Db_details,text='Host name',padx=3,font=('Arial',10,'bold'))
Host.grid(row=4,column=0)
Entry3=Entry(Db_details,bd=5)
Entry3.grid(row=4,column=1)

service=Label(Db_details,text='Service name',padx=3,font=('Arial',10,'bold'))
service.grid(row=5,column=0)
Entry4=Entry(Db_details,bd=5)
Entry4.grid(row=5,column=1)

connect=Button(Db_details,text='Connect',pady=5,font=('Arial',10,'bold'),command=connect)
connect.grid(row=6,column=2,sticky=W+E)

list_label=Label(main_frame,text='List of Components',font=('Arial',10,'bold'))
list_label.grid(row=7,column=0,sticky=W,pady=5)
#===============TEXT FRAME======================================================
text_frame=Frame(root,height=10,width=100,borderwidth=5)
text_frame.grid(row=2,column=0,sticky=W+E)

scrollbar = Scrollbar(text_frame)
scrollbar.grid(row=1,column=5)
scrollbar.fill=Y

#Creating the Label for Text widget
l1=Label(text_frame,text='Results')
l1.grid(row=0,column=0,sticky=W)

#Creating the Text widget
Textbox = Text(text_frame,width=90,height=10)
Textbox.grid(row=1,columnspan=4,pady=4)
Textbox.tag_configure('big', font=('Verdana', 10, 'bold'))

# attach listbox to scrollbar
Textbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=Textbox)

#=========================READING CONTENTS FROM SHHET AND CREATING FRAMES,BUTTONS AND LABELS==============================================

l=0
with open('readdata.csv','r+') as f:
  write=csv.reader(f)
  header=next(write)
  for n,i in enumerate(write):
           if i[0] and i[1] and  not i[2]:
               l +=1
               listofmodules.append(i[0])
               module_name.append([i[1]])
print (listofmodules)
print (module_name)
#=====================================================
frames = []
j=0
col=0
#color=['blue','red','Green']
for i in range(1,len(listofmodules)+1,1):
    frame = Frame(root, borderwidth=1
                  #bg = color[col]
                  )
    frame.grid(row = 0, column = 0, sticky = W+E+N+S)
    frames.append(frame)
    j +=1
    col +=1
#==========================================================
#Function to raise the mainframe
def go_to_main_frame():
    Textbox.delete(1.0,END)
    text_frame.grid_forget()
    main_frame.tkraise()

    
    

    
#==================================listofmodules====================
def case(a,b,label,label1,label2,testcase,sql):
    print (label)
    print (label1)
    print (label2)
    print (sql)
    print (a)
    print (b)
    Textbox.delete(1.0,END)

    cursor = connection.cursor()
    cursor.execute(sql)
    for i in cursor:
         dict1[testcase]='FAIL'
         b.set(value=1)
         Textbox.insert(1.0,i)
         Textbox.insert(END, '\n')
    if cursor.rowcount == 0:
            a.set(value=1)
            Textbox.insert(1.0,'All looks good')
            dict1[testcase]='PASS'
    cursor.close()
    
##########################################################
def hiding_frame(frame,frameinfo):
    frame.tkraise()
    text_frame.grid(row=1,column=0,sticky=W+E)
    print (frame)
    print (frameinfo + 'frame has been hidden')

###############################################################
list_pass=[]
list_fail=[]
r=0
c=0
ro=20
with open('readdata.csv','r+') as f:
      write=csv.reader(f)
      header=next(write)
      for n,i in enumerate(write):
            for d in range(0,len(listofmodules),1):
                
                if str(i[0])== str(listofmodules[d])  and i[1] and not i[2]:
                  label=Button(main_frame,text=i[1],state=DISABLED,command=partial(hiding_frame,frames[d],str(i[1])))
                  label.grid(row=ro,column=0,sticky=W)
                  dict1[i[1]]=''
                  ro=ro+1
                if str(i[0])== str(listofmodules[d])  and i[1] and  i[2]:
                  label=Button(frames[d],text=i[1])
                  label.grid(row=r,column=1,padx=5,sticky=W)
                  vars()['pass'+str(i[1])]=IntVar()
                  label1=Checkbutton(frames[d],text='Pass', variable=vars()['pass'+str(i[1])])
                  label1.grid(row=r,column=2,padx=5,sticky=W)
                  vars()['fail'+str(i[1])]=IntVar()
                  label2=Checkbutton(frames[d], text='Fail', variable=vars()['fail'+str(i[1])])
                  label2.grid(row=r,column=3,padx=5,sticky=W)
                  label3=Button(frames[d],text=val,command=partial(case,vars()['pass'+str(i[1])],
                                                                   vars()['fail'+str(i[1])],
                                                                   str(label.cget('text')),
                                                                   str(label1.cget('text')),
                                                                   str(label2.cget('text')),
                                                                   str(i[1]),
                                                                   str(i[3])))
                  label3.grid(row=r,column=4,padx=5)
                  dict1[i[1]]='Not Run'
                  #print (frames[d])
                  r=r+1
#Creating Back button in main_frame
for i in range(0,len(listofmodules),1):
    back = Button (frames[i],text='Back',command=go_to_main_frame)
    back.grid(row=30,column=1,sticky=W,ipadx=2,pady=5)
    listofbackbuttons.append(back)

##########################################################################
def report():
    global count
    global filename
    filename='Test Execution Report'+ str(count)+'.csv'
    keyList = dict1.keys()
    valueList = dict1.values()

    rows = zip(keyList, valueList)

    with open(filename , 'w+') as f:
        
       writer = csv.writer(f)
       for row in rows:
         writer.writerow(row)
    count=count+1
    #=========Sending Mail============================================
def send_mail():
    import smtplib
    import mimetypes
    from email.mime.multipart import MIMEMultipart
    from email import encoders
    from email.message import Message
    from email.mime.audio import MIMEAudio
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.text import MIMEText

    emailfrom = "balaji.ma@prodapt.com"
    emailto = 'balaji.ma@prodapt.com'
    fileToSend = filename
    username = "balaji.ma"
    password = "Bal@j199"

    msg = MIMEMultipart()
    msg["From"] = emailfrom
    msg["To"] = emailto
    msg["Subject"] = "Test case execution summary Report"
    msg.preamble = "Test case execution summary Report"

    ctype, encoding = mimetypes.guess_type(fileToSend)
    if ctype is None or encoding is not None:
     ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
     fp = open(fileToSend)
    # Note: we should handle calculating the charset
     attachment = MIMEText(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "image":
     fp = open(fileToSend, "rb")
     attachment = MIMEImage(fp.read(), _subtype=subtype)
     fp.close()
    elif maintype == "audio":
     fp = open(fileToSend, "rb")
     attachment = MIMEAudio(fp.read(), _subtype=subtype)
     fp.close()
    else:
     fp = open(fileToSend, "rb")
     attachment = MIMEBase(maintype, subtype)
     attachment.set_payload(fp.read())
     fp.close()
     encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)
    msg.attach(attachment)

    server = smtplib.SMTP("outlook.prodapt.com:587")
    server.starttls()
    server.login(username,password)
    server.sendmail(emailfrom, emailto, msg.as_string())
    server.quit()
export=Button(main_frame,text='Export',command=report)
export.grid(row=40,column=8,sticky=W,padx=10)
mail=Button(main_frame,text='Send Mail',command=send_mail)
mail.grid(row=40,column=5,sticky=W)
#########################################################################
#connect()	
#frames[0].grid_forget()

#frames[1].tkraise()
#frames[0].tkraise()
#frames[2].tkraise()
#frames[1].grid_forget()
#frames[0].grid_forget()
#frames[2].grid_forget()
text_frame.grid_forget()
main_frame.grid_forget()
login_frame.tkraise()


root.mainloop()
