from tkinter import *
import pandas as pd
import numpy as np
import smtplib
import matplotlib
import xlrd
import xlwt
from xlutils.copy import copy
from tkinter import messagebox
from tkinter.simpledialog import askfloat
from tkinter.simpledialog import askstring
from tkinter.simpledialog import askinteger
import matplotlib.pyplot as plt
from threading import Thread
from re import fullmatch
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pretty_html_table import build_table
from PIL import ImageTk,Image
import tkinter.font as font



def load():
        global book
        book=xlrd.open_workbook("C:your file location")
        global sheet
        sheet=book.sheet_by_index(0)
        global wb
        wb=copy(book)
        global wsheet
        wsheet=wb.get_sheet(0)
        global mails
        mails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
def save():
    wb.save('your file location')
def startsession():
    global s
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login('youremail@gmail.com', "your app generated password from google")
    print("Logged In")
class main(Tk):
    def __init__(self):
        super().__init__()
        self.title("Marks Management System")
        self.geometry("700x700")
        img = ImageTk.PhotoImage(Image.open("home.jpg"))
        label = Label(self.master, image=img)
        label.image = img 
        label.place(x=0, y=0, relwidth=1, relheight=1)
        self.resizable(0,0)
        button_font = font.Font(family="Century")
        Button(self,text="Student Wise Results",font=button_font,command=lambda:self.stuwres(self),bg='#45b592',fg='#ffffff',activebackground="lightblue",height=5,width=20).place(x=100,y=40)
        Button(self,text="Subject Wise Results",font=button_font,command=lambda:self.subwres(self),bg='#45b592',fg='#ffffff',activebackground="lightblue",height=5,width=20).place(x=400,y=40)
        Button(self,text="Subject Wise Comparision",font=button_font,command=lambda:self.subwcom(self),bg='#45b592',fg='#ffffff',activebackground="lightblue",height=5,width=20).place(x=100,y=220)
        Button(self,text="Send all reports to Students",font=button_font,command=self.sendreports,bg='#45b592',fg='#ffffff',activebackground="lightblue",height=5,width=20).place(x=400,y=220)
        Button(self,text="Send weak student reports",font=button_font,height=5,width=20,bg='#45b592',fg='#ffffff',activebackground="lightblue",command=self.sweaksreports).place(x=100,y=400)
        Button(self,text="Add or Update Marks",font=button_font,height=5,width=20,bg='#45b592',fg='#ffffff',activebackground="lightblue",command=lambda:self.addmarks(self)).place(x=400,y=400)
        Button(self,text="Send Report to Teachers",font=button_font,height=5,width=20,bg='#45b592',fg='#ffffff',activebackground="lightblue",command=self.sreptchrs).place(x=100,y=580)
        Button(self,text="EXIT",height=5,font=button_font,width=20,bg="red",activebackground="#f24438",fg="white",activeforeground="White",command=lambda:self.destroy()).place(x=400,y=580)
    def sreptchrs(self):
        startsession()
        load()
        def runthread():
            sheet2 = book.sheet_by_index(1)
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            marks = []
            emails = [sheet2.cell_value(k,1) for k in range(1,sheet2.nrows)]
            for i in range(len(emails)):
                marks = [sheet.cell_value(j,i+2) for j in range(1,sheet.nrows)]
                df = pd.DataFrame(data={"Student Name":names,"Marks":marks})
                html = '''
                        <html>
                            <body>
                                <h1>Report of {0}</h1>
                                  {1}
                                    </body>
                                </html>
                                '''.format(sheet.cell_value(0,i+2),build_table(df, 'blue_light'))
                email_message = MIMEMultipart()
                email_message['From'] = "youremail@gmail.com"
                email_message['To'] = emails[i]
                email_message['Subject'] = "Report Of Your Subject"
                email_message.attach(MIMEText(html, "html"))
                email_string = email_message.as_string()
                s.sendmail("youremail@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Sent To Teachers")
        s.quit()
    def sweaksreports(self):
        startsession()
        load()
        def runthread():
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            inde = ['Maths','Physics','Chemistry','Biology','English','Hindi','Total','Percentage']
            emails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
            for i in range(len(names)):
                if(int(sheet.cell_value(i+1,9))<60):
                    marks = [sheet.cell_value(i+1,j+2) for j in range(8)]

                    df = pd.DataFrame(data={"Subject":inde,"Marks":marks})
                    html = '''
                                    <html>
                                        <body>
                                            <h1>Report of {0}</h1>
                                            <h2>Your ward is a weak student</h2>
                                            {1}
                                        </body>
                                    </html>
                                    '''.format(names[i],build_table(df, 'blue_light'))
                    email_message = MIMEMultipart()
                    email_message['From'] = "youremail@gmail.com"
                    email_message['To'] = emails[i]
                    email_message['Subject'] = "Report Card"
                    email_message.attach(MIMEText(html, "html"))
                    email_string = email_message.as_string()
                    s.sendmail("youremail@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Cards of Weak Students Sent Successfully")
        s.quit()
    def subwcom(self,s):
        x1 = [2,4,6,8,10,12]
        y1 = []
        load()
        for j in range(6):
            y1.append(np.mean(np.array([sheet.cell_value(i+1,j+2) for i in range(sheet.nrows-1)])))
        plt.bar(x1,y1,tick_label=['Maths','Physics','Chemistry','Biology','English','Hindi'],color='green',width=0.8)
        plt.plot(x1,[35 for i in range(6)],linestyle='dashed',label='Fail',color='red')
        plt.plot(x1,[60 for i in range(6)],linestyle='dashed',color='orange',label='weak')
        plt.ylim(0,100)
        plt.xlabel('Subjects')
        plt.ylabel('Mean Marks')
        plt.title('Subject Wise Comparison')
        for j in range(6):
            plt.text(x1[j],y1[j],'%.1f'%y1[j])
        plt.legend()
        plt.show()
    def sendreports(self):
        startsession()
        load()
        def runthread():
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            inde = ['Maths','Physics','Chemistry','Biology','English','Hindi','Total','Percentage']
            emails = [sheet.cell_value(k,1) for k in range(1,sheet.nrows)]
            for i in range(len(names)):
                marks = [sheet.cell_value(i+1,j+2) for j in range(8)]

                df = pd.DataFrame(data={"Subject":inde,"Marks":marks})
                html = '''
                                <html>
                                    <body>
                                        <h1>Report of {0}</h1>
                                        {1}
                                    </body>
                                </html>
                                '''.format(names[i],build_table(df, 'blue_light'))
                email_message = MIMEMultipart()
                email_message['From'] = "youremail@gmail.com"
                email_message['To'] = emails[i]
                email_message['Subject'] = "Report Card"
                email_message.attach(MIMEText(html, "html"))
                email_string = email_message.as_string()
                s.sendmail("youremail@gmail.com",emails[i],email_string)
        t1=Thread(target=runthread)
        t1.start()
        t1.join()
        messagebox.showinfo("Success","Report Cards Sent Successfully")
        s.quit()
    def subwres(self,s):
        s.withdraw()
        f=Toplevel(s)
        f.grab_set()
        f.geometry("400x500")
        img = ImageTk.PhotoImage(Image.open("400.jpg"))
        label = Label(f, image=img)
        label.image = img 
        label.place(x=0, y=0, relwidth=1, relheight=1)

        f.resizable(0,0)
        f.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))
        def subwmar():
            if(sub.get()==0):
                return
            load()
            names = [sheet.cell_value(k,0) for k in range(1,sheet.nrows)]
            x1=[k*2 for k in range(1,len(names)+1)]
            marks = [sheet.cell_value(k,int(sub.get())) for k in range(1,sheet.nrows)]
            plt.bar(x1,marks,tick_label=names,color='green',width=0.8)
            plt.plot(x1,[35 for i in range(len(names))],linestyle='dashed',label='Fail',color='red')
            plt.plot(x1,[60 for i in range(len(names))],linestyle='dashed',color='orange',label='weak')
            plt.ylim(0,100)
            plt.xlabel('Students')
            plt.ylabel('Marks')
            plt.title('Results of %s'%sheet.cell_value(0,int(sub.get())))
            for j in range(len(names)):
                plt.text(x1[j],marks[j],'%.1f'%marks[j])
            plt.legend()
            plt.show()
        Label(f,text="Select Subject To Display",bg='#45b592',fg='#000000',font=("century",16)).place(anchor=CENTER,x=200,y=50)
        sub=IntVar()
        Radiobutton(f,text="Maths",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=2).place(x=140,y=90)
        Radiobutton(f,text="Physics",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=3).place(x=140,y=130)
        Radiobutton(f,text="Chemistry",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=4).place(x=140,y=170)        
        Radiobutton(f,text="Biology",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=5).place(x=140,y=210)
        Radiobutton(f,text="English",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=6).place(x=140,y=250)
        Radiobutton(f,text="Hindi",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=7).place(x=140,y=290)
        Button(f,text="Submit",bg="Green",fg="White",font=("century",12),activebackground="Green",command=subwmar).place(x=80,y=350)
        Button(f,text="Cancel",bg="Red",fg="White",font=("century",12),activebackground="Red",command=lambda:(f.destroy(),s.deiconify())).place(x=255,y=350)
        
    def stuwres(self,s):
        s.withdraw()
        email = askstring("E-Mail", "Enter Your Mail : ")
        if(email!=None):
            load()
            if(email in mails):
                r = mails.index(email)
                x1=[2,4,6,8,10,12,14]
                y1=[sheet.cell_value(r+1,i) for i in range(2,8)]
                y1.append(sheet.cell_value(r+1,9))
                plt.bar(x1,y1,tick_label=['Maths','Physics','Chemistry','Biology','English','Hindi','Percentage'],color='green',width=0.8)
                plt.plot(x1,[35 for i in range(7)],linestyle='dashed',label='Fail',color='red')
                plt.plot(x1,[60 for i in range(7)],linestyle='dashed',color='orange',label='weak')
                plt.ylim(0,100)
                plt.xlabel('Subjects')
                plt.ylabel('Marks')
                plt.title('Results of %s'%sheet.cell_value(r+1,0))
                for j in range(7):
                    plt.text(x1[j],y1[j],'%.1f'%y1[j])
                plt.legend()
                plt.show()
            else:
                messagebox.showinfo("Invalid", "No Such Mail Exist !!")
        s.deiconify()
    def addmarks(self,s):
        s.withdraw()
        f=Toplevel(s)
        def update():
            load()
            m =  askstring("E-Mail", "Enter Student's Mail : ")
            if(m!=None):
                if(m in mails):
                    up = Toplevel(f)
                    def updsub():
                        newmark = askfloat("Update Marks", "Enter New Marks ",parent=up)
                        if(newmark!=None):
                            if(newmark>100 or newmark<0):
                                messagebox.showwarning("Invalid","Marks range is 0-100")
                                return
                            wsheet.write(r,int(sub.get()),float(newmark))
                            t = np.sum(np.array([float(sheet.cell_value(r,i)) for i in range(2,8)]))
                            wsheet.write(r,8,t)
                            wsheet.write(r,9,t/6)
                            save()
                            messagebox.showinfo("Saved","Marks Updated Successfully")
                    up.grab_set()
                    up.geometry("400x500")
                    img = ImageTk.PhotoImage(Image.open("400.jpg"))
                    label = Label(up, image=img)
                    label.image = img 
                    label.place(x=0, y=0, relwidth=1, relheight=1)
                    up.resizable(0,0)
                    up.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))
                    r = mails.index(m)+1
                    Label(up,text="Select Subject To Update",bg='#45b592',fg='#000000',font=("century",16)).place(anchor=CENTER,x=200,y=50)
                    sub=IntVar()
                    Radiobutton(up,text="Maths",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=2).place(x=140,y=90)
                    Radiobutton(up,text="Physics",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=3).place(x=140,y=130)
                    Radiobutton(up,text="Chemistry",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=4).place(x=140,y=170)
                    Radiobutton(up,text="Biology",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=5).place(x=140,y=210)
                    Radiobutton(up,text="English",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=6).place(x=140,y=250)
                    Radiobutton(up,text="Hindi",bg='lightblue',fg='#000000',activebackground="orange",anchor="w",font=("century",12),width=9,variable=sub,value=7).place(x=140,y=290)
                    Button(up,text="Submit",font=("century",12),bg="Green",fg="White",activebackground="Green",command=updsub).place(x=80,y=350)
                    Button(up,text="Cancel",font=("century",12),bg="Red",fg="White",activebackground="Red",command=lambda:up.destroy()).place(x=255,y=350)
                else:
                    messagebox.showinfo("Invalid","No Such Mail Exist !!")
            
        def cupdate():
            try:
                vals = np.array([float(m.get()),float(p.get()),float(c.get()),float(b.get()),float(e.get()),float(h.get())])
            except ValueError:
                messagebox.showwarning("Mandatory", "All Feilds are Mandatory !\nMarks should be integers",parent=f)
                return
            if(name.get()=='' or email.get()==''):
                messagebox.showwarning("Mandatory","Email and Name Feilds are Mandatory",parent=f)
                return
            if(np.any(vals>100) or np.any(vals<0)):
                messagebox.showwarning("Invalid","Marks range is 0-100",parent=f)
                return
            if(not fullmatch(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', email.get())):
                messagebox.showinfo("Invalid mail", "Enter a Valid Mail")
                return
            if(not messagebox.askyesno("Confirmation","Are you sure to Save ? ")):
                return
            load()
            r = sheet.nrows
            wsheet.write(r,0,name.get())
            wsheet.write(r,1,email.get())
            wsheet.write(r,2,float(m.get()))
            wsheet.write(r,3,float(p.get()))
            wsheet.write(r,4,float(c.get()))
            wsheet.write(r,5,float(b.get()))
            wsheet.write(r,6,float(e.get()))
            wsheet.write(r,7,float(h.get()))
            t = np.sum(np.array([float(m.get()),float(p.get()),float(c.get()),float(b.get()),float(e.get()),float(h.get())]))
            wsheet.write(r,8,t)
            wsheet.write(r,9,t/6)
            save()
            messagebox.showinfo("Success","Marks Added Successfully!")
            name.delete(0,END)
            p.delete(0,END)
            m.delete(0,END)
            c.delete(0,END)
            b.delete(0,END)
            e.delete(0,END)
            h.delete(0,END)
            email.delete(0,END)
        f.grab_set()
        f.geometry("400x500")
        img = ImageTk.PhotoImage(Image.open("400.jpg"))
        label = Label(f, image=img)
        label.image = img 
        label.place(x=0, y=0, relwidth=1, relheight=1)
        f.resizable(0,0)
        f.protocol("WM_DELETE_WINDOW",lambda:(s.deiconify(), f.destroy()))
        Label(f,text="Enter the Details",bg='#45b592',fg='#000000',font=("century",16)).place(anchor=CENTER,x=200,y=35)
        Label(f,text="Enter Your Name : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=75)
        name=Entry(f,width=30,bg='wheat')
        name.place(x=200,y=75)
        Label(f,text="Enter your Email : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=105)
        email=Entry(f,width=30,bg='wheat')
        email.place(x=200,y=105)
        Label(f,text="Maths Marks : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=160)
        m=Entry(f,width=10,bg='wheat')
        m.place(x=200,y=160)
        Label(f,text="Physics Marks : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=190)
        p=Entry(f,width=10,bg='wheat')
        p.place(x=200,y=190)
        Label(f,text="Chemistry Marks : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=220)
        c=Entry(f,width=10,bg='wheat')
        c.place(x=200,y=220)
        Label(f,text="Biology Marks : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=250)
        b=Entry(f,width=10,bg='wheat')
        b.place(x=200,y=250)
        Label(f,text="English Marks : ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=280)
        e=Entry(f,width=10,bg='wheat')
        e.place(x=200,y=280)
        Label(f,text="Hindi Marks :  ",bg='lightblue',fg='#000000',anchor="w",font=("century",12),width=15).place(x=20,y=310)
        h=Entry(f,width=10,bg='wheat')
        h.place(x=200,y=310) 
        Button(f,text="Submit",bg="Green",fg="White",font=("century",12),activebackground="Green",command=cupdate).place(x=80,y=420)
        Button(f,text="Cancel",bg="Red",fg="White",activebackground="Red",font=("century",12),command=lambda:(f.destroy(),s.deiconify())).place(x=255,y=420)
        Button(f,text="Update Marks",bg="orange",activebackground="Yellow",command=update,font=("century",12)).place(x=140,y=365)
main().mainloop()