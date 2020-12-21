from tkinter import *
import csv
from tkinter.filedialog import asksaveasfile
from tkinter import messagebox
from tkinter import filedialog
import sqlite3
import datetime
from tkcalendar import *
import os
import pandas as pd
import xlrd
from datetime import date
from babel.dates import format_date, parse_date, get_day_names, get_month_names
from babel.numbers import *
from PIL import ImageTk,Image
import webbrowser

file_address=None
mark_p=[]


#_________________________Functions________________________________________



def close(y):
    close_value=messagebox.askokcancel('GMeet Attendance','Sure To Close')
    if close_value==1:
        if y=='login_w':
            login_w.destroy()
        elif y=='top':
            top.destroy()
        elif y=='Reset_window':
            Reset_window.destroy()
        else:
            pass
def logout():
    close_value=messagebox.askokcancel('GMeet Attendance','Sure To Logout')
    if close_value==1:
        top.destroy()
    else:
        pass
def remove(x):
    x.place(x=-500,y=-500)
def login_page():
    try:
        try:
            conn=sqlite3.connect('icon\data.db')
            db=conn.cursor()
            db.execute('CREATE TABlE IF NOT EXISTS LOGIN (id integer PRIMARY KEY,username text,password text,Recovery_key text,subject text,signup_date date)')
            conn.commit()
            db.close()
        except Exception  as e:
            #print(e)
            pass
        finally:

                
            #____________Removing Elements_________________________________
            remove(Create_new)
            remove(username_text)
            remove(password1_text)
            remove(again_password)
            remove(username_enter)
            remove(password1)
            remove(again_password)
            remove(Recovery_key_text)
            remove(Recovery_key)
            remove(C1)
            remove(C2)
            remove(login_b_2)
            remove(Reset_password_b)
            remove(Close)
            remove(signup_b)
            remove(again_password_text)

            remove(Student_Names_text)
            remove(Student_names_Add_show)
            remove(Student_names_upload)
            remove(Subject_name_text)
            remove(Subject_name)
        
    
    except Exception:
        pass
    
    finally:
        global login_icon
        global username
        global password
        global login_button
        global signup_button
        login_icon=Label(login_w,image=login_img,bg='#000000')
        login_icon.place(x=280,y=110)

        username=Entry(login_w,width=25,borderwidth=7,bg='#000000',fg="#4db8ff",font=' kalam 14 ')
        username.place(x=200,y=220)
        username.insert(0,'Username')


        password=Entry(login_w,width=25,borderwidth=7,bg='#000000',fg="#4db8ff",font=' kalam 14 ',show='*')
        password.place(x=200,y=270)
        password.insert(0,'Password')
        

        login_button=Button(login_w,image=login_button_img,borderwidth=0,highlightthickness = 0,command=login)
        login_button.place(x=350,y=320)

        signup_button=Button(login_w,image=signup_img,borderwidth=0,highlightthickness = 0,command=signup_page)
        signup_button.place(x=190,y=320)
        
def login():
    conn=sqlite3.connect('icon\data.db')
    db=conn.cursor()
    if (len(username.get())==0):
        messagebox.showerror('Login','Enter The Username')
    elif (len(password.get())==0):
        messagebox.showerror('Login','Enter The Password')
    else:
        db.execute('''SELECT * FROM LOGIN WHERE username='{}' '''.format(username.get().upper()))
        Value_1=db.fetchall()
        if Value_1:

            if (Value_1[0][2])==password.get():
                messagebox.showinfo('Login','Successfully Login')
                GMeet_Attendance_W(username.get().upper())
            else:
                messagebox.showerror('Login','Incorrect Password')
        else:
            messagebox.showerror('Login','User Not Found')
    db.close()


def upload_students():
    login_w.path__=filedialog.askopenfilename(initialdir="",title="Select Student List",filetypes=(("Excel File",".xlsx"),("All files",".*")))
    global Student_xl_file
    Student_xl_file=login_w.path__
    Student_names_Add_show.delete(0,END)
    Student_names_Add_show.insert(0,Student_xl_file)

def reset_password_fun(u,rk,p,ap):
    if len(u)==0:
        messagebox.showerror('GMeet Attendance','Please Enter Username')
    elif len(rk)==0:
        messagebox.showerror('GMeet Attendance','Please Enter Recovery Key')
    elif len(p)==0:
        messagebox.showerror('GMeet Attendance','Please Enter Password')
    elif len(ap)==0:
        messagebox.showerror('GMeet Attendnace','Please Enter Again Password')
    elif p!=ap:
        messagebox.showerror('GMeet Attendance','Password Not Matched')
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    db.execute("""SELECT * FROM LOGIN WHERE username=='{}' """.format(u.upper()))
    all_data=db.fetchall()
    conn.commit()
    
    if all_data:
        if all_data[0][3]!=rk:
            messagebox.showerror('GMeet Attendance','Invalid Recovery Key')
        else:
            db.execute("UPDATE LOGIN SET password='{}' WHERE username='{}'".format(p,u.upper()))
            conn.commit()
            messagebox.showinfo('GMeet Attendance','Password Successfully Changed')
    else:
        messagebox.showerror('GMeet Attendance','Username Not found')
    db.close()
        
def clos():
    global Resert_window
    Reset_window.destroy()
            
def reset_password_page():
    global Reset_window
    global icon
    Reset_window=Toplevel()
    Reset_window.geometry('650x650+0+0')
    Reset_window.resizable(width=False,height=False)
    Reset_window.iconbitmap(icon)
    Reset_window.configure(background="#000000")
    Reset_window.title("GMeet Attendance")

    head=Label(Reset_window,text="GMeet Attendance ",bg='#000000',fg="#7C4521",font=' kalam 29 bold',pady=10)
    Reset_password_text=Label(Reset_window,text="Reset Password ",bg='#000000',fg='red',font=' kalam 28 bold',pady=8)
    username_text1=Label(Reset_window,text="Username :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    Recovery_key_text=Label(Reset_window,text="Recovery Key :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)

    username_entered=Entry(Reset_window,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 14 ')
    Recovery_key_entered=Entry(Reset_window,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 14 ',show='*')

    NewPassword_1_text=Label(Reset_window,text="New Password : ",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    NewPassword1=Entry(Reset_window,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 14 ',show='*')

    NewPassword_2_text=Label(Reset_window,text="Again New Password : ",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    NewPassword2=Entry(Reset_window,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 14 ',show='*')

    close_Reset=Button(Reset_window,text='Close',width=10,font='kalam 10 bold',fg='#040BFC',command=clos)
    Reset_button=Button(Reset_window,text='Reset',width=10,font=' kalam 10 bold',fg='#040BFC',command=lambda:reset_password_fun(username_entered.get(),Recovery_key_entered.get(),NewPassword1.get(),NewPassword2.get()))
    

    head.place(x=1,y=-5)
    Reset_password_text.place(x=240,y=50)
    username_text1.place(x=140,y=120)
    Recovery_key_text.place(x=120,y=170)
    username_entered.place(x=270,y=127)
    Recovery_key_entered.place(x=270,y=177)

    NewPassword1.place(x=270,y=240)  
    NewPassword_1_text.place(x=110,y=230)

    NewPassword_2_text.place(x=55,y=300)
    NewPassword2.place(x=270,y=300)

    
    Reset_button.place(x=200,y=400)

    
    close_Reset.place(x=340,y=400)



    Reset_window.mainloop()




def signup_page():
    
    #_____________________Removing Elements
    remove(login_icon)
    remove(username)
    remove(password)
    remove(login_button)
    remove(signup_button)
    #messagebox.showinfo('GMeet Attendance','Account Successfully Created ')
    

    #____________________Adding Elements

    global Create_new
    global username_text
    global password1_text
    global username_enter
    global password1
    global again_password
    global again_password_text
    global Recovery_key_text
    global Recovery_key
    global C1
    global C2
    global CheckVar1
    global login_b_2
    global Reset_password_b
    global Close
    global signup_b
    
    global Student_Names_text
    global Student_names_Add_show
    global Student_names_upload
    global Subject_name_text
    global Subject_name
    

    Create_new=Label(login_w,text="Create New Account ",bg='#000000',fg="#4db8ff",font=' kalam 29 bold',pady=10)
    Create_new.place(x=180,y=1)
    
    username_text=Label(login_w,text="Username :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    username_text.place(x=80,y=80)
    
    password1_text=Label(login_w,text="Password :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    password1_text.place(x=80,y=130)
    
    again_password_text=Label(login_w,text="Again Password :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    again_password_text.place(x=80,y=180)
    
    username_enter=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ')
    username_enter.place(x=280,y=80)
    
    password1=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ',show='*')
    password1.place(x=280,y=130)
    
   
    
    again_password=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ',show='*')
    again_password.place(x=280,y=180)



    Recovery_key_text=Label(login_w,text="Recovery Key :",bg='#000000',fg='#FC6604',font=' kalam 16 bold',pady=8)
    Recovery_key_text.place(x=80,y=250)
    
    Recovery_key=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ',show='*')
    Recovery_key.place(x=280,y=250)



    Subject_name_text=Label(login_w,text="Subject Name :",bg='#000000',fg="#FC6604",font=' kalam 16 bold',pady=10)
    Subject_name=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ')

    Subject_name_text.place(x=80,y=305)
    Subject_name.place(x=280,y=305)
    
    
    Student_Names_text=Label(login_w,text="Students Name with Enroll :",bg='#000000',fg="#FC6604",font=' kalam 14 bold',pady=10)
    Student_names_Add_show=Entry(login_w,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 16 ')
    Student_names_upload=Button(login_w,text='Upload',width=10,font=' kalam 10 bold',fg='#FC6604',command=upload_students)

    Student_Names_text.place(x=25,y=360)
    Student_names_Add_show.place(x=280,y=360)
    Student_names_upload.place(x=280,y=410)
    



    CheckVar1 = IntVar()
    C1 = Checkbutton(login_w, text = "I have installed Meet Attendance Extension in my Chrome Browser", variable = CheckVar1, onvalue = 1, offvalue = 0, bg='#000000',fg="#FC6604")
    C2 = Checkbutton(login_w, text = "I Agree to All T&C", variable = CheckVar1,onvalue = 1, offvalue = 0,bg='#000000',fg="#FC6604")

    C1.place(x=240,y=470)
    C2.place(x=240,y=500)
    

    login_b_2=Button(login_w,text='Login',width=10,font=' kalam 10 bold',fg='#040BFC',command=login_page)
    login_b_2.place(x=80,y=550)
    
    
    Reset_password_b=Button(login_w,text='Reset Password ? ',width=20,font=' kalam 10 bold',fg='#040BFC',command=reset_password_page)
    Reset_password_b.place(x=190,y=550)
    
    Close=Button(login_w,text='Close' ,width=10,font=' kalam 10 bold',fg='#040BFC',command=lambda:close('login_w'))
    Close.place(x=360,y=550)
    
    signup_b=Button(login_w,text='SignUp',width=10,font='kalam 10 bold',fg='#040BFC',command=sign_up)
    signup_b.place(x=470,y=550)



def sql_cmd_create(path_,u):
    
    wb=xlrd.open_workbook(path_)
    sheet=wb.sheet_by_index(0)
    y=int(sheet.nrows)
    temp_cmd="CREATE  TABLE {} (id integer PRIMARY KEY, date date, time time, "
    for i in range(1,y):
        tenroll=sheet.cell_value(i,2).upper()
        temp_cmd=temp_cmd+'\"'+tenroll+'\"'+' text'
        if i!=y-1:
            temp_cmd=temp_cmd+','
    temp_cmd=temp_cmd + ')'
    return temp_cmd.format(u.upper())


def sign_up():
    conn=sqlite3.connect('icon\data.db')
    db=conn.cursor()    
    if len(username_enter.get())==0:
        messagebox.showerror('SignUp','Please Enter Username')
    elif len(password1.get())==0:
        messagebox.showerror('SignUp','Please Enter Password')
    elif len(again_password.get())==0:
        messagebox.showerror('SignUp','Please Enter Password Again')
    elif len(Recovery_key.get())==0:
        messagebox.showerror('SignUp','Please Enter Recovery Key')
    elif CheckVar1.get()==0:
        messagebox.showerror('Signup','Check The Checkboxes')
    elif len(Subject_name.get())==0:
        messagebox.showerror('Signup','Please Enter Subject Name')
        
    elif len(Student_names_Add_show.get())==0:
        messagebox.showerror('Signup','Please Upload The Excel File which contains Students Names and Enrollment')
     

    else:
            db.execute("Select * from LOGIN where username='{}'".format(username_enter.get().upper()))
            value=db.fetchall()
            conn.commit()   
            if value:
                messagebox.showerror('Login','User Already Exist')

            elif password1.get()!=again_password.get():
                messagebox.showerror('Login','Password Not Match')
            else:

                #____________________________________Creating DB TaBle For Student Name__________________________________________________

                try:
                    u=username_enter.get()
                    u=u+'_D'
                    db.execute("CREATE TABLE '{}' (id integer PRIMARY KEY, NAMES text, ENROLL text)".format(u.upper()))
                    conn.commit()
                    
                #____________________________________SAving Excel Details in DB File________________________________________________________
                
                    wb=xlrd.open_workbook(Student_xl_file)
                    sheet=wb.sheet_by_index(0)
                    y=int(sheet.nrows)
                    x=int(sheet.ncols)
                    for i in range(1,y):
                        fname=sheet.cell_value(i,0)
                        lname=sheet.cell_value(i,1)
                        tenroll=sheet.cell_value(i,2)
                        name_=fname.upper()+lname.upper()
                        db.execute('''INSERT INTO '{}' (NAMES, ENROLL ) VALUES ('{}','{}')'''.format(u,name_,tenroll.upper()))
                        conn.commit()

            #_______________________________________Saving Details For login _________________________________________________________________
                
                    dt=date.today()
                    #print('Install Date save on : ',dt)
                    db.execute("insert into LOGIN (username,password ,Recovery_key,subject,signup_date ) VALUES ('{}','{}','{}','{}','{}')".format(username_enter.get().upper(),password1.get(),Recovery_key.get(),Subject_name.get(),dt))
                    conn.commit()
                    messagebox.showinfo("Signup","Successfully Signup")    
                    #print('Saving Login Details get executed')
                #______________________Creating New Table For New User___________________________________________________________
                

                    
                    db.execute(sql_cmd_create(Student_xl_file,username_enter.get().upper()))
                    #print(sql_cmd_create(Student_xl_file,username_enter.get().upper()))
                    conn.commit()
                    login_page()
                    #print('A New Table Created')
                    
                except Exception as ckmb:
                    pkmb=str(ckmb)
                    #print(pkmb)
                    if pkmb=="""name 'Student_xl_file' is not defined""":
                        del_temp_cmd="""DROP TABLE {};""".format(u.upper())
                        db.execute(del_temp_cmd)
                        conn.commit()
                        #print("Deleted")
                        messagebox.showerror('GMeet Attendance','Please Upload a Valid Excel file which contains Students Names and Enrollment')
                    elif pkmb=="""list index out of range""":
                        del_temp_cmd="""DROP TABLE {};""".format(u.upper())
                        db.execute(del_temp_cmd)
                        conn.commit()
                        #print("Deleted")
                        messagebox.showerror('Error','Error Found In Excel File')
                        
    db.close()



def name_to_enroll(username,name):
    try:
        u=username.upper()+'_D'
        con=sqlite3.connect('icon/data.db')
        db=con.cursor()
        db.execute(''' SELECT * FROM '{}' WHERE NAMES =='{}' '''.format(u,name))
        con.commit()
        kk=db.fetchall()
        #print('121 fetched')
        #print(kk[0])
        
        con.close()
        return kk[0][2].upper()
    except Exception as e:
        ee=str(e)
        #print(e)
        if ee=='list index out of range':
            messagebox.showerror('Error','Unable to found Enrollment for {}'.format(name))
            return None

def Modify_Student_page(u):
    global MSP
    MSP=Tk()
    MSP.geometry('650x650+0+0')
    MSP.resizable(width=False,height=False)
    MSP.configure(background="#000000")
    MSP.title("GMeet Attendance")
    global icon
    MSP.iconbitmap(icon)
    title_lb=Label(MSP,text="Modify Student Detalis",bg='#000000',fg="#FC6604",font=' kalam 18 bold',pady=0)
    title_line=Label(MSP,text="___________________________________________\n___________________________________________",bg='#000000',fg="#FC6604",font='times 11 bold',pady=0)
    title_lb.place(x=10,y=10)
    title_line.place(x=10,y=40)
    #________________________Defining Texts____________________________
    previuos_fname_tx=Label(MSP,text='Previous First Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    previous_lname_tx=Label(MSP,text='Previous Last Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    previous_enroll_tx=Label(MSP,text='Previous Enrollment No. : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    previous_fname=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    previous_lname=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    previous_enroll=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    

    new_fname_tx=Label(MSP,text='New First Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    new_lname_tx=Label(MSP,text='New Last Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    new_enroll_tx=Label(MSP,text='New Enrollment No. : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    new_fname=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    new_lname=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    new_enroll=Entry(MSP,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')



    Modify_btn=Button(MSP,text='Modify',width=20,command=lambda:Modify_Student_fun(u.upper(),previous_fname.get().upper(),previous_lname.get().upper(),previous_enroll.get().upper(),new_fname.get().upper(),new_lname.get().upper(),new_enroll.get().upper()))

    #att_head=Label(MSP,text='Attention Here ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)

    #__________________________Placing Them_______________________________
    previuos_fname_tx.place(x=50,y=130)
    previous_fname.place(x=300,y=130)
    previous_lname_tx.place(x=50,y=180)
    previous_lname.place(x=300,y=180)
    previous_enroll_tx.place(x=50,y=230)
    previous_enroll.place(x=300,y=230)

    new_fname_tx.place(x=50,y=280)
    new_fname.place(x=300,y=280)
    new_lname_tx.place(x=50,y=330)
    new_lname.place(x=300,y=330)
    new_enroll_tx.place(x=50,y=380)
    new_enroll.place(x=300,y=380)

    Modify_btn.place(x=290,y=440)
    MSP.mainloop()
    

def data_modify(u,nfn,nln,nen,idn):
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    uu=u+'_D'
    cmd_2="UPDATE '{}' SET NAMES='{}',ENROLL='{}' WHERE id=='{}' ".format(uu,nfn+nln,nen,idn)
    db.execute(cmd_2)
    conn.commit()
    #print('Updated')
    
    
def Modify_Student_fun(u,pfn,pln,pen,nfn,nln,nen):
    check_var=messagebox.askokcancel('GMeet Attendance','Sure To Edit Details')
    if check_var==1:
        if len(pfn)==0:
            messagebox.showerror('GMeet Attendance','Please Enter Previous First Name')
        elif len(pln)==0:
            messagebox.showerror('GMeet Attendance','Please Enter Previous Last Name')
        elif len(pen)==0:
            messagebox.showerror('GMeet Attendance','Please Enter Previous Enrollment Number')
        elif len(nfn)==0:
            messagebox.showerror('GMeet Attendance','Please Enter New First Name')
        elif len(nln)==0:
            messagebox.showerror('GMeet Attendance','Please Enter New Last Name')
        elif len(nen)==0:
            messagebox.showerror('GMeet Attendance','Please Enter New Enrollment Number ')
        else:
            conn=sqlite3.connect('icon/data.db')
            db=conn.cursor()
            try:
                cmd_1="SELECT * FROM '{}' WHERE ENROLL =='{}' ".format(u+'_D',nen)
                #print(cmd_1)
                db.execute(cmd_1)
                conn.commit()
                x_data=db.fetchall()
                #print(x_data)
                conn.commit()
                if len(x_data)!=0:
                    pass
                else:
                    
                    cmd_4=''' SELECT * FROM '{}' WHERE ENROLL=='{}' '''.format(u+'_D',pen)
                    db.execute(cmd_4)
                    conn.commit()
                    xl_data=db.fetchall()
                    #print(xl_data)
                    #print(xl_data[0][1])
                    #print(xl_data[0][2])
                    if (xl_data[0][1])!=(pfn+pln):
                        messagebox.showerror('Error','Previous Name not matched')
                        #print('Your Previous Name Not Matched')
                    elif xl_data[0][2]!=pen:
                        messagebox.showerror('Error','Previous Enrollment Not Matched')
                        #print('Your Previous Enrollment Not Matched')
                    else: 
                        cmd_3='''ALTER TABLE '{}'  RENAME COLUMN '{}' TO '{}' '''.format(u,pen,nen)
                        #print(cmd_3)
                        db.execute(cmd_3)
                        conn.commit()
                        idn=xl_data[0][0]
                        data_modify(u,nfn,nln,nen,idn)
                        messagebox.showinfo('GMeet Attendance','Successfully Details are Edited')
                        global MSP
                        MSP.destroy()
                        
                        
            except Exception as e:
                #print(e)
                qq=str(e)
                if qq=="list index out of range":
                    messagebox.showerror('Error','Enrollment No. not matched with DB')
                    pass
                else:
                    
                    messagebox.showerror('Error','Details are not Matched, Please Check Again')
                    pass
            db.close()
    else:
        pass
            
def student_list_fun(u):
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    file=pd.read_sql_query(''' SELECT * FROM '{}' ;'''.format(u.upper()+'_D'),conn)
        
    files = [ ('Excel Files', '*.xlsx')] 
    f = asksaveasfile(title='Save Student List',mode='a',filetypes = files, defaultextension = files)
    if f==None:
        pass
    else:
        f.close()
        os.remove(f.name)
        #print(f.name)
        file.to_excel(f.name)
        messagebox.showinfo('GMeet Attendance','Successfully Saved Student List')
        #print('Successfully Saved')
        db.close()
def add_student_page(u):
    global AS
    AS=Toplevel()
    AS.geometry('650x650+0+0')
    AS.resizable(width=False,height=False)
    global icon
    AS.iconbitmap(icon)
    AS.configure(background="#000000")
    AS.title("GMeet Attendance")
    title_lb=Label(AS,text="Add Student ",bg='#000000',fg="#FC6604",font=' kalam 18 bold',pady=0)
    title_line=Label(AS,text="___________________________________________\n___________________________________________",bg='#000000',fg="#FC6604",font='times 11 bold',pady=0)
    title_lb.place(x=10,y=10)
    title_line.place(x=10,y=40)
    #________________________Defining Texts____________________________
    s_fname_tx=Label(AS,text='Student First Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    s_lname_tx=Label(AS,text='Student Last Name : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    s_enroll_tx=Label(AS,text='Student Enrollment No. : ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)
    s_fname=Entry(AS,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    s_lname=Entry(AS,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    s_enroll=Entry(AS,borderwidth=4,bg='#000000',fg="#FC6604",font=' times 16 bold')
    

    del_btn=Button(AS,text='ADD',width=20,command=lambda:add_student_fun(u.upper(),s_fname.get().upper(),s_lname.get().upper(),s_enroll.get().upper()))

    #att_head=Label(MSP,text='Attention Here ',bg='#000000',fg="#FC6604",font=' times 16 bold',pady=0)

    #__________________________Placing Them_______________________________
    s_fname_tx.place(x=50,y=130)
    s_fname.place(x=300,y=130)
    s_lname_tx.place(x=50,y=180)
    s_lname.place(x=300,y=180)
    s_enroll_tx.place(x=50,y=230)
    s_enroll.place(x=300,y=230)

    del_btn.place(x=290,y=280)
    AS.mainloop()


def sql_add_col(u,fn,ln,en):    
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    
    sql_query3=''' INSERT INTO '{}' (NAMES, ENROLL ) VALUES ('{}','{}')'''.format(u+'_D',fn+ln,en)
    db.execute(sql_query3)
    conn.commit()
    sql_query2=''' ALTER TABLE '{}' ADD '{}' text; '''.format(u,en)
    db.execute(sql_query2)
    conn.commit()
    db.close()
    #print(sql_query2)
    #print(sql_query3)

def add_student_fun(u,fn,ln,en):
    check_var=messagebox.askokcancel('GMeet Attendance','Sure To Add ')
    if check_var==1:
        if len(fn)==0:
            messagebox.showerror('GMeet Attendance','Please Enter First Name')
        elif len(ln)==0:
            messagebox.showerror('GMeet Attendance','Please Enter Last Name')
        elif len(en)==0:
            messagebox.showerror('GMeet Attendance','Please Enter Enrollment Number')
        else:
            conn=sqlite3.connect('icon/data.db')
            db=conn.cursor()
            cmd_check="SELECT * FROM '{}' WHERE ENROLL=='{}' ".format(u+'_D',en)
            conn.commit()
            db.execute(cmd_check)
            all_data=db.fetchall()
            conn.commit()
            #print(all_data)
            if len(all_data)==0:
                
                sql_add_col(u,fn,ln,en)
                messagebox.showinfo('Success','{} Added Successfully'.format(fn+' '+ln))
                global AS
                AS.destroy()
                

            else:
                messagebox.showerror('Error','Enrollment already Exist Kindley  Delete it first')
            db.close()
    else:
         pass
def support_us(u):
    global SU
    #Qr_img=ImageTk.PhotoImage(Image.open("icon/qr.png"))
    SU=Tk()
    SU.geometry('650x650+0+0')
    SU.resizable(width=False,height=False)
    SU.configure(background='#000000')
    global icon
    SU.iconbitmap(icon)
    
    SU.title("Support Us - GMeet Attendance")
    head2=Label(SU,text="Support Us ",bg='#000000',fg='red',font=' Pacifico 24 bold',pady=15)
    head2.place(x=10,y=20)
    head3=Label(SU,text="Love Our Work ?? \n ",bg='#000000',fg='red',font=' Acme 16 bold',pady=15)
    l4=Label(SU,text="if no then let us know the reason  \n ",bg='#000000',fg='red',font=' Acme 13 bold')
    head3.place(x=190,y=130)
    l4.place(x=130,y=180)
    l5=Label(SU,text='''if Yes then you can support us by contributing a little amount
             from which we can add addition feature on it & also provide \n more software like it in future''',bg='#000000',fg='green',font=' Acme 13 bold')
    l5.place(x=30,y=220)
    m1=Label(SU,text="UPI Id :jagdishpal02000-2@okaxis  \n ",bg='#000000',fg='red',font=' Acme 13 bold')
    m1.place(x=150,y=299)
    #qr=Label(SU,text='qr',bg='#000000',font=' times 16 bold',fg="#ffffff",pady=8)
    #Qr=Label(SU,image=Qr_img)
    #Qr.place(x=120,y=420)
    last=Label(SU,text="Your Smallest contribution is worth for us  \n Thank You {}  :) \n love from GMeet Attendance".format(u.upper()),bg='#000000',fg='red',font=' Acme 13 bold')
    last.place(x=120,y=340)
    #________________________________________________Functions : ON_________________________________________________________________________

def get_su_date(u):
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    date_cmd=''' SELECT * FROM LOGIN WHERE USERNAME =='{}' '''.format(u)
    db.execute(date_cmd)
    d_value=db.fetchall()
    return d_value[0][5]

def view_attendance(user_name):
    global vwat
    vwat=Tk()
    vwat.geometry('650x650+0+0')
    vwat.resizable(width=False,height=False)
    vwat.configure(background='#000000')
    global icon
    vwat.iconbitmap(icon)
    vwat.title("View Attendance - GMeet Attendance")

    head_tx=Label(vwat,text='View  Attendance',bg='#000000',font=' times 25 bold',fg="#00ff80",pady=8)
    head_tx.place(x=180,y=20)
    to_text=Label(vwat,text=" To ",bg='#000000',fg="#ff9900",font=' kalam 29 bold',pady=10)
    from_text=Label(vwat,text=" From ",bg='#000000',fg="#ff9900",font=' kalam 29 bold',pady=10)
    from_text.place(x=45,y=120)
    to_text.place(x=400,y=120)
    now=datetime.datetime.now()
    Day=now.day
    Month=now.month
    Year=now.year
    #print('to',Day,Month,Year)
    kk=get_su_date(user_name.upper())
    now_f=datetime.datetime.strptime(kk,'%Y-%m-%d')
    #print(now)

    Dayf=now_f.day
    Monthf=now_f.month
    Yearf=now_f.year
    #print('From data : ',Dayf,Monthf,Yearf)
    global from_d
    from_d=DateEntry(vwat,width=30,bg="#000000",fg="#ff9900",selectmode="day",day=Dayf,month=Monthf,year=Yearf)
    from_d.place(x=20,y=220)
    
    global to_d
    to_d=DateEntry(vwat,width=30,bg="#000000",fg="#ff9900",selectmode="day",day=Day,month=Month,year=Year)
    to_d.place(x=350,y=220)
    t_a = IntVar()
    t_v= Checkbutton(vwat, text ="Download Total Attendance", font='times 18 bold',variable = t_a ,onvalue = 1, offvalue = 0,bg='#000000',fg="#FC6604")
    t_v.place(x=215,y=280)
    save_btn=Button(vwat,text='Save Attendance',command=lambda:save_at(user_name,t_a.get()))
    save_btn.place(x=270,y=340)
    vwat.mainloop()


def save_at(user_name,t_A):
    x=from_d.get_date()
    y=to_d.get_date()
    fd=datetime.datetime.strftime(x,'%Y-%m-%d')
    td=datetime.datetime.strftime(y,'%Y-%m-%d')
        
    #print(fd)
    #print(td)
    conn=sqlite3.connect('icon/data.db')
    #print('sucessfully Connnected')
    #print(t_A)

    if t_A==1:
        file=pd.read_sql_query(''' SELECT * FROM '{}';'''.format(user_name),conn)
    else: 
            file=pd.read_sql_query(''' SELECT * FROM '{}'
                                    WHERE date BETWEEN '{}' AND '{}';'''.format(user_name,fd,td),conn)
    
    files = [ ('Excel Files', '*.xlsx')] 
    f = asksaveasfile(title='Save Attendance',mode='a',filetypes = files, defaultextension = files)
    if f==None:
        pass
    else:
        f.close()
        os.remove(f.name)
        #print(f.name)
        file.to_excel(f.name)
        messagebox.showinfo('GMeet Attendance','Successfully Saved Attendance')
        #print('Successfully Saved')
        vwat.destroy()
        
    

def upload():
    top.path_=filedialog.askopenfilename(initialdir="",title="Select A E xcel File",filetypes=(("CSV File",".csv"),("All files",".*")))
    global file_address
    global upload_add
    file_address=top.path_
    upload_add.delete(0,END)
    upload_add.insert(0,file_address)

def return_name(a):
    for i in range(0,len(a)):
        if a[i]=='?':
            index1=i
            break
    temp=a[0:index1]       
    p=temp.replace('\t','')
    q=p.replace(' ','')
    return q.upper()
            





def sql_cmd_insert(u,date,time):
    
    conn=sqlite3.connect('icon/data.db')
    db=conn.cursor()
    db.execute("SELECT * FROM '{}' ".format(u+'_D'))
    exl_data=db.fetchall()
    conn.commit()
    d_=str(date)
    t_=str(time)
    y=len(exl_data)
    temp_cmd="INSERT INTO {} (date,time,"
    for i in range(0,y):
        tenroll=exl_data[i][2].upper()
        temp_cmd=temp_cmd+'\"'+tenroll+'\"'
        #print(exl_data[i][1])
        if i!=y-1:
            temp_cmd=temp_cmd+','
    temp_cmd=temp_cmd + ') VALUES (' +'\"'+d_+'\"'+',' + '\"'+t_+'\"'+','

    for i in range(0,y):
        temp_cmd=temp_cmd+'\'A\' '
        if i!=y-1:
            temp_cmd=temp_cmd+','

    temp_cmd=temp_cmd+')'
    
    db.close()
    return temp_cmd.format(u.upper())



def initialization(user_name):
    conn=sqlite3.connect('icon\data.db')
    db=conn.cursor()
    global date
    global time
    date=datetime.datetime.now().date()
    time=datetime.datetime.now().time()

    #print(date,time)
    initialization_cmd=sql_cmd_insert(user_name,date,time)
    #print(initialization_cmd)
    db.execute(initialization_cmd)
    conn.commit()
    #print('Data Initialized Succesfully')
    db.close()

def total_time(a):
    time=''
    flag=1
    for i in range(0,len(a)):
        if a[i]=='(':
            i1=i
        if a[i]==')':
            i2=i
            flag=0
        if flag==0:  
            time=time+a[i1+1:i2]
            flag=1
    #for total time
    t1=time.replace('min', "+")
    tt=t1[0:len(t1)-1]
    #print('Total Time For Every One is : ',tt)
    return eval(tt)



def save_attendance(user_name):
    global mark_p
    global file_address
    global upload_add
    global m_time
    global En_Err
    global nma
    nma=0
    if file_address!=None:
        #k=m_time.get().replace("Min","")
        #min_time=k[0:len(k)-1]
        conn=sqlite3.connect('icon\data.db')
        db=conn.cursor()
        #db.execute('''SELECT * FROM '{}' '''.format(user_name))
        #all_data=db.fetchall()
        ##print(all_data)
        #conn.commit()
        try:
            with open(file_address,'r') as file:
                csv_reader=csv.reader(file)
                #print(csv_reader)
                next(file)
                next(file)
                next(file)
                initialization(user_name)
                for i in csv_reader:
                    
                    name=return_name(i[0])
                    #print(name,len(name))
                    Enroll=name_to_enroll(user_name,name)
                    #print('Enroll - ',Enroll)
                    #t_t=str(total_time(i[0]))
                    ##print('Total Time',t_t)
                    ##print('Min Time',min_time)
                    #print(i[0])
                    no_enroll_f=str(Enroll)
                    #print(no_enroll_f[0])
                    if Enroll == None:
                        En_Err.insert(0.0,'\n {}'.format(name))
                        nma=nma+1
                        
                        #print('Not Mark Att of ',nma)
                    else:
                        #if(eval(t_t)>=eval(min_time)):
                        attendance_save_command='''UPDATE '{}'
                                                    SET '{}'='P'
                                                    WHERE date='{}' and time='{}';'''.format(user_name,Enroll,date,time)
                        db.execute(attendance_save_command)
                        conn.commit()
                        #print('Attendance of {} is added'.format(Enroll))
                        #print(name)
                        #else:
                         #   #print('{} is apsent '.format(Enroll))
                          #  Mark_as_absent.insert(0.0,'{} Marked Absent \n'.format(name))
                           # mark_p.append(Enroll)
                            
                # En_Err.insert(0,'''\n '{}' Attendance is not marked Becauese Unable to find their Enrollments kindley  Check The Student List'''.format(nma))    
            
            messagebox.showinfo('GMeet Attendance','Successfully Saved')
            upload_add.delete(0,END)
            file_address=None
            if nma!=0:
                En_Err.insert(0.0,'''\n '{}' Attendance is  not marked Becauese Unable to find their Enrollments kindley  Check The Student   List :-'''.format(nma))
                    
            ##print(mark_p)
            global top
            #tet='''These Students are Mark \n Absent Due to Time Quanta '''
            #Mark_as_Absent_tx=Label(top,text=tet,bg='black',fg='white')
            #Mark_as_Absent_tx.place(x=485,y=260)
            db.close()
        except Exception as g:
            #print(g)
            upload_add.delete(0,END)
            file_address=None
            
            gg=str(g)
            #print(gg,type(gg))

            #if gg=='unexpected EOF while parsing (<string>, line 0)':
             #       if nma!=0:
              #          En_Err.insert(0.0,'''\n '{}' Attendance is not marked Becauese Unable to find their Enrollments Kindaly Check The Student List'''.format(nma))
                    
            db.close()

                
    else:
        messagebox.showerror('GMeet Attendance','Please Upload CSV File')
        


    #-------------------------------------------Function : OFF---------------------------------------------------------------------------------

    


def GMeet_Attendance_W(user_name):
    global login_w
    login_w.destroy()
    global top
    top=Tk()
    top.geometry('650x650+0+0')
    top.resizable(width=False,height=False)
    top.configure(background='#000000')
    top.title("GMeet Attendance")
    global icon
    icon="icon/logo.ico"
    top.iconbitmap(icon)

        #_________________Main Menu : on_______________

    main_menu=Menu()
    Account=Menu(main_menu,tearoff=False)
    Edit=Menu(main_menu,tearoff=False)
    View=Menu(main_menu,tearoff=False)
    Help=Menu(main_menu,tearoff=False)
    About=Menu(main_menu,tearoff=False)
    support=Menu(main_menu,tearoff=False)


    Account.add_command(label='Change Password',command=reset_password_page)
    Account.add_command(label='Logout',command=logout)

    Edit.add_command(label='Add Student',command=lambda:add_student_page(user_name))
    Edit.add_command(label='Modify Student',command=lambda:Modify_Student_page(user_name))

    View.add_command(label='View Attendance',command=lambda:view_attendance(user_name))
    View.add_command(label='Student List',command=lambda:student_list_fun(user_name))

    Help.add_command(label='Read Article',command=website)
    Help.add_command(label='Watch Video',command=video)
    Help.add_command(label='Other Help',command=video)

    About.add_command(label='About App',command=AboutApp)
    About.add_command(label='About Us',command=about_page)

    support.add_command(label='Support Us',command=lambda:support_us(user_name))
    #Cascading

    main_menu.add_cascade(label='Account',menu=Account)
    main_menu.add_cascade(label='Edit',menu=Edit)
    main_menu.add_cascade(label='View',menu=View)
    main_menu.add_cascade(label='Help',menu=Help)
    main_menu.add_cascade(label='About',menu=About)
    main_menu.add_cascade(label='Support Us',menu=support)
    top.config(menu=main_menu)



    #-----------------Main Menu : off--------------------------
        



    #___________________________________________Home Screen : ON________________________________________________________________________________

    heading=Label(top,text='GMeet Attendance',bg='#000000',fg="#7C4521",font=' kalam 29 bold',pady=10)
    heading.place(x=10,y=5)


    csv_icon=PhotoImage(file="icon/csv_f.png")
    submit_b=PhotoImage(file="icon/save_attendance.png")
    csv_file_icon=Label(top,image=csv_icon,bg='#000000',font=' times 16 bold',fg="#ffffff",pady=8)
    csv_file_icon.place(x=180,y=120)
    global upload_add
    upload_add=Entry(top,width=25,borderwidth=5,bg='#000000',fg="#FC6604",font=' kalam 14 ')
    upload_add.place(x=245,y=130)


    #time_sel_txt=Label(top,text='Minimum Time For Marking Present',bg='#000000',font=' kalam 11 bold',fg="#ffffff",pady=8)
    #time_sel_txt.place(x=100,y=250)
    #_______________________________________Time Menu Option
    #global m_time
    #m_time=StringVar()
    #m_time.set("15 Min")
    #drop=OptionMenu(top,m_time,"0 Min","5 Min","10 Min","15 Min","20 Min","30 Min","40 Min","60 Min","75 Min")
    #drop.place(x=390,y=250)
    #drop.config(bg='black',fg='#1AFC07')
    upload_button=Button(top,text='Upload CSV File',bg='black',fg='#07FCFC',command=upload)
    upload_button.place(x=270,y=180)
    



    submit=Button(top,image=submit_b,borderwidth=0,highlightthickness = 0,command=lambda:save_attendance(user_name))
    submit.place(x=240,y=325)
    global Mark_as_absent
    global En_Err
    En_Err=Text(top,width=20,height=18,bg='#000000',fg='white')

    En_Err.place(x=10,y=300)
    #Mark_as_absent=Text(top,width=20,height=18,bg='#000000',fg='white')
    #Mark_as_absent.place(x=480,y=300)
    #save_present_anyway=Button(top,text='Save Present Anyway',bg='black',fg='#FC2107',command=lambda:Save_anyway(user_name))
    #save_present_anyway.place(x=500,y=603)
    clear_names=Button(top,text='Clear Names',command=Cl_names,bg='black',fg='#FC2107')
    clear_names.place(x=30,y=603)
    top.mainloop()
    



    #--------------------------------------Home Screen : OFF -----------------------------------------------------------------------------------------

def video():
    webbrowser.open('https://youtu.be/_OMdJFEiuGY')

def website():
    messagebox.showerror('GMeet_Attendnace','This is Beta version website will update soon')

def Cl_names():
    global En_Err
    global nma
    try:
        if nma!=0:
            En_Err.delete(0.0,END)
        else:
           messagebox.showerror('Error',' Sorry List is Empty :( ')
    except NameError:
        messagebox.showerror('Error',' Sorry List is Empty :( ')

def Save_anyway(u):
    if len(mark_p)!=0:
        global Mark_as_absent
        conn=sqlite3.connect('icon/data.db')
        db=conn.cursor()
        for i in range(0,len(mark_p)):
            #print(mark_p[i])
            Enroll=mark_p[i]
            #print(u,Enroll)
            attendance_save_command='''UPDATE '{}' SET '{}'='P'
                                    WHERE date='{}' and time='{}';'''.format(u,Enroll,date,time)
            db.execute(attendance_save_command)
            conn.commit()
            #print('Attendance of {} is added'.format(Enroll))
        Mark_as_absent.delete(0.0,END)
        messagebox.showinfo('GMeet Attendance','Successfully Saved as Present')
    else:
        messagebox.showerror('Error','Sorry List is EmPty :( ' )

def AboutApp():
    AA=Tk()
    AA.geometry('650x650+0+0')
    AA.resizable(width=False,height=False)
    AA.configure(background='#000000')
    global icon
    AA.iconbitmap(icon)
    AA.title("About App - GMeet Attendance")
    h_label=Label(AA,text='GMeet Attendance',bg='#000000',font=' Pacifico 18 bold',fg="#dd05fa",pady=8)
    tx= """ This app is made for the Teachers
            who are using Google Meet for online class, they can use it
            just need to install Meet Attendance in chrome extension in the browser and
            use GMeet Attendance(our software) to Manage and save the attendance. """

    metter=Label(AA,text=tx,bg='black',fg='#0dfcec',font=' Acme 14  ')

    b1=Button(AA,text='Visit Website',width=10,font='kalam 10 bold',fg='#040BFC')
    b2=Button(AA,text='Watch Video',width=10,font='kalam 10 bold',fg='#040BFC')
    b3=Button(AA,text='Read Article',width=10,font='kalam 10 bold',fg='#040BFC')
    h_label.place(x=210,y=10)
    b1.place(x=180,y=250)
    b2.place(x=280,y=250)
    b3.place(x=380,y=250)
    metter.place(x=2,y=90)
    AA.mainloop()

def about_page():
    Reset_window=Toplevel()
    Reset_window.geometry('650x650+0+0')
    Reset_window.resizable(width=False,height=False)
    Reset_window.configure(background='#000000')
    global icon
    Reset_window.iconbitmap(icon)
    Reset_window.title("About us - GMeet Attendance")
    head2=Label(Reset_window,text="About Us",bg='#000000',fg='red',font=' Pacifico 24 bold',pady=15)
    xx='''
        This Software is made to save and manage the Attendance from Google Meet.
        The developer of GMeet Attendnace are : \n
                Jagdish Pal             Nishant Tiwari          Akash Deep Suman
                \n    Student of Rewa Engineering College Rewa M.P.
                if you need any kind of help related to software or want a  
                Software, you just need to contact us. We are here to help

                Mobile :  +91 7024964179,
                             +91 8770183159,
                             +91 9131095491
                
                Email :     jagdishpal02000@gmail.com
                            nnisshaantt@gmail.com
                            deep98akash@gmail.com

                website : hallwick.in

        Want to send Result from Gmail to all your Students, You can use Result Mailer 
        '''
    body=Label(Reset_window,text=xx,bg='#000000',fg="#0dfcec",font=' Acme 13 bold',pady=8)
    visit_website=Button(Reset_window,text='Visit Website',font='kalam 10 bold',fg='#040BFC',command=website)
    visit_website.place(x=170,y=530)
    thanks=Label(Reset_window,text='Thank You :) ',bg='black',fg="#dd05fa",font=' kalam 25 bold',pady=8)
    Result_mailer=Button(Reset_window,text='Download Result Mailer',font='kalam 10 bold',fg='#040BFC',command=rm)
    Result_mailer.place(x=325,y=530)
    thanks.place(x=10,y=575)

    head2.place(x=210,y=1)
    body.place(x=0,y=75)
    Reset_window.mainloop()
def rm():
    webbrowser.open('https://drive.google.com/file/d/1G67Ic8MfhjW5RP1ARhA7XfyzOjjAltCw/view')

#------------------------------------------------------------Functions : OFF---------------------------------------------------------------------------------------






login_w=Tk()
#_________________________________________________________Basic Settings of window__________________________________________________________________________________
login_w.geometry('650x650+0+0')
login_w.resizable(width=False,height=False)
login_w.configure(background='#000000')
login_w.title("GMeet Attendance")
global icon
icon="icon/logo.ico"
login_w.iconbitmap(icon)

login_img=PhotoImage(file="icon/login.png")
login_button_img=PhotoImage(file="icon/login_button.png")
signup_img=PhotoImage(file="icon/signup.png")

#_________________________________________________________HomePage ___________________________________________________________________________________


login_page()
login_w.mainloop()
