from tkinter import*
from PIL import ImageTk
from tkinter import messagebox, filedialog
import os
import pandas as pd
import email_function
import time

class BUlk_EMAIL:
    def __init__ (self,root):
        self.root=root
        self.root.title("BULK EMAIL SENDER APPLICATION")
        self.root.geometry("1000x550+200+50")
        self.root.resizable(False,False)
        self.root.config(bg="white")
        self.root.focus_force()

        #=======Icons============
        self.email_icon=ImageTk.PhotoImage(file="images/email.png")
        self.setting_icon=ImageTk.PhotoImage(file="images/setting.png")

        #----------Title--------------
        title=Label(self.root,text="Bulk Email Send Panel",image=self.email_icon,padx=10,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        decs=Label(self.root,text="Use Excel File to Send the Bulk Email at once,with just one click. Ensure the Email Column Name must be Email.",font=("Calibri (Body)",14),bg="#FFD966",fg="#262626").place(x=0,y=80,relwidth=1)

        btn_setting=Button(self.root,image=self.setting_icon,bd=0,activebackground="#222A35",bg="#222A35",cursor="hand2",command=self.setting_window).place(x=880,y=5)

        #-----------------------------
        self.var_choice=StringVar()
        self.var_choice.set("single")
        single=Radiobutton(self.root,value="single",variable=self.var_choice,text="Single",font=("times new roman",30,"bold"),bg="white",fg="#262626",activebackground="white",command=self.check_single_or_bulk).place(x=50,y=150)
        bulk=Radiobutton(self.root,value="bulk",variable=self.var_choice,text="Bulk",font=("times new roman",30,"bold"),bg="white",fg="#262626",activebackground="white",command=self.check_single_or_bulk).place(x=250,y=150)
        

        #-----------------------------------
        to=Label(self.root,text="TO (Email Address):",font=("times new roman",18),bg="white").place(x=50,y=250)
        subj=Label(self.root,text="SUBJECT:",font=("times new roman",18),bg="white").place(x=50,y=300)
        msg=Label(self.root,text="MESSAGE:",font=("times new roman",18),bg="white").place(x=50,y=350)

        self.txt_to=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_to.place(x=300,y=250,width=350,height=30)

        self.btn_browse=Button(self.root,text="BROWSE",font=("times new roman",15,"bold"),bg="#8FAADC",fg="#262626",cursor="hand2",activebackground="#8FAADC",activeforeground="#262626",command=self.browse_files)
        self.btn_browse.place(x=670,y=250,height=30,width=120)

        self.txt_subj=Entry(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_subj.place(x=300,y=300,width=450,height=30)

        self.txt_msg=Text(self.root,font=("times new roman",14),bg="lightyellow")
        self.txt_msg.place(x=300,y=350,width=650,height=120)

        #---------------Status-------------------
        self.lbl_total=Label(self.root,font=("times new roman",18),bg="white")
        self.lbl_total.place(x=50,y=490)

        self.lbl_sent=Label(self.root,font=("times new roman",18),bg="white",fg="green")
        self.lbl_sent.place(x=300,y=490)

        self.lbl_left=Label(self.root,font=("times new roman",18),bg="white",fg="orange")
        self.lbl_left.place(x=420,y=490)

        self.lbl_failed=Label(self.root,font=("times new roman",18),bg="white",fg="red")
        self.lbl_failed.place(x=550,y=490)
        

        btn_clear=Button(self.root,text="CLEAR",font=("times new roman",18,"bold"),bg="#262626",fg="white",cursor="hand2",activebackground="#262626",activeforeground="white",command=self.clear1).place(x=700,y=490,height=30,width=120)
        btn_send=Button(self.root,text="SEND",font=("times new roman",18,"bold"),bg="#00B0F0",fg="white",cursor="hand2",activebackground="#00B0F0",activeforeground="white",command=self.send_email).place(x=830,y=490,height=30,width=120)
        self.btn_browse.config(state=DISABLED)
        self.check_file_exist()

    def setting_window(self):
        self.check_file_exist()
        self.root2=Toplevel()
        self.root2.title("Settings")
        self.root2.geometry("700x310+350+90")
        self.root2.config(bg="white")
        self.root2.resizable(False,False)
        self.root2.focus_force()
        self.root2.grab_set()

        title2=Label(self.root2,text="Credentials Settings",image=self.setting_icon,padx=10,compound=LEFT,font=("Goudy Old Style",48,"bold"),bg="#222A35",fg="white",anchor="w").place(x=0,y=0,relwidth=1)
        decs2=Label(self.root2,text="Enter the Email address and password from which to send all emails.",font=("Calibri (Body)",14),bg="#FFD966",fg="#262626").place(x=0,y=80,relwidth=1)

        from_=Label(self.root2,text="Email Address:",font=("times new roman",18),bg="white").place(x=50,y=150)
        pass_=Label(self.root2,text="Password:",font=("times new roman",18),bg="white").place(x=50,y=200)

        self.txt_from=Entry(self.root2,font=("times new roman",14),bg="lightyellow")
        self.txt_from.place(x=300,y=150,width=350,height=30)

        self.txt_pass=Entry(self.root2,font=("times new roman",14),bg="lightyellow",show="•")
        self.txt_pass.place(x=300,y=200,width=350,height=30)
        
        btn_clear2=Button(self.root2,text="CLEAR",font=("times new roman",18,"bold"),bg="#262626",fg="white",cursor="hand2",activebackground="#262626",activeforeground="white",command=self.clear2).place(x=300,y=260,height=30,width=120)
        btn_save=Button(self.root2,text="SAVE",font=("times new roman",18,"bold"),bg="#00B0F0",fg="white",cursor="hand2",activebackground="#00B0F0",activeforeground="white",command=self.save_setting).place(x=430,y=260,height=30,width=120)

        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)


    def send_email(self):
        x=len(self.txt_msg.get('1.0',END))
        if self.txt_to.get()=="" or self.txt_subj.get()=="" or x==1:
            messagebox.showerror('Error','All Fields Are Required',parent=self.root)
        else:
            if self.var_choice.get() == 'single':
                status=email_function.email_send_funct(self.txt_to.get(),self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                if status == 's':
                    messagebox.showinfo('Success','Email Has Been Sent Successfully')
                if status == 'f':
                    messagebox.showerror('Failed','Email Not Sent,Please Try again!!')
            if self.var_choice.get() == 'bulk':
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send_funct(x,self.txt_subj.get(),self.txt_msg.get('1.0',END),self.from_,self.pass_)
                    if status =='s':
                        self.s_count+=1
                    if status =='f':
                        self.f_count+=1
                    self.status_bar()
                    time.sleep(1)
                messagebox.showinfo('Success','Email Has Been Sent Successfully.Please Check Staus!!')


    def check_single_or_bulk(self):
        if self.var_choice.get()=='single':
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)

        if self.var_choice.get()=='bulk':
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state='readonly')

    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_subj.delete(0,END)
        self.txt_msg.delete('1.0',END)
        self.var_choice.set('single')
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text='')
        self.lbl_sent.config(text='')
        self.lbl_left.config(text='')
        self.lbl_failed.config(text='')

    def save_setting(self):
        if self.txt_from.get()=="" or self.txt_pass.get()=="":
            messagebox.showerror('Error','All Fields Are Required',parent=self.root2)
        else:
            f=open('important.txt','w')
            f.write(self.txt_from.get()+","+self.txt_pass.get())
            f.close()
            messagebox.showinfo('Success','Credentials Saved Successfully')
            self.check_file_exist()

    def check_file_exist(self):
        if os.path.exists('important.txt')==False:
            f=open('important.txt','w')
            f.write(",")
            f.close()
            
        f2=open('important.txt','r')
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(",")[0],i.split(",")[1]])
        #print(self.credentials)
        self.from_ = self.credentials[0][0]
        self.pass_ = self.credentials[0][1]
        #print(self.from_,self.pass_)

    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)

    def browse_files(self):
        op=filedialog.askopenfile(initialdir='/',title="Select Excel File For Emails",filetypes=(("Alll Files","*.*"),("Excel Files",".xlsx")))
        if op != None:
            data=pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.emails=list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split('/')[-1]))
                    self.txt_to.config(state='readonly')
                    self.lbl_total.config(text='Total:'+str(len(self.emails)))
                    self.lbl_sent.config(text='SENT:')
                    self.lbl_left.config(text='LEFT')
                    self.lbl_failed.config(text='FAILED:')
                else:
                    messagebox.showerror('Error',"This Files Do not Have Any Emails.",parent=self.root)
            else:
                messagebox.showerror('Error','Please Select File Which Have Email Columns',parent=self.root)

    def status_bar(self):
        self.lbl_total.config(text='STATUS:'+str(len(self.emails))+"=>>")
        self.lbl_sent.config(text='SENT:'+str(self.s_count))
        self.lbl_left.config(text='LEFT:'+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text='FAILED:'+str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()





root=Tk()
obj=BUlk_EMAIL(root)
root.mainloop()