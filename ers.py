#Employee Record System 
from tkinter import*
from tkinter import messagebox
from openpyxl import load_workbook
import xlrd
import pandas as pd


global username_verify
global password_verify

global username_login_entry
global password_login_entry


root=Tk()                               #Main window 
f=Frame(root)
frame1=Frame(root)
frame2=Frame(root)
frame3=Frame(root)
loginFrame = Frame(root,bg="white",width=300,height=450)
root.title("Employee Record System")
root.attributes('-fullscreen', True)
#root.geometry("650x650")
root.configure(background="Grey")

scrollbar=Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)

firstname=StringVar()                    #Declaration of all variables
lastname=StringVar()
id=StringVar()
dept=StringVar()
designation=StringVar()
remove_firstname=StringVar()
remove_lastname=StringVar()
searchfirstname=StringVar()
searchlastname=StringVar()
sheet_data=[]
row_data=[]





def emp_dict(*args):                   #To add a new entry and check if entry already exist in excel sheet
    #print("done")
    workbook_name="sample.xlsx"
    workbook=xlrd.open_workbook(workbook_name)
    worksheet=workbook.sheet_by_index(0)
    
    wb=load_workbook(workbook_name)
    page=wb.active
    
    p=0
    for i in range(worksheet.nrows):
        for j in range(worksheet.ncols):
            cellvalue=worksheet.cell_value(i,j)
            print(cellvalue)   
            sheet_data.append([])
            sheet_data[p]=cellvalue
            p+=1
    print(sheet_data)
    fl=firstname.get()
    fsl=fl.lower()
    ll=lastname.get()
    lsl=ll.lower()
    if (fsl and lsl) in sheet_data:
        print("found")
        messagebox.showerror("Error","This Employee already exist")
    else:
        print("not found")
        for info in args:
            page.append(info)
        messagebox.showinfo("Done","Successfully added the employee record")

    wb.save(filename=workbook_name)
    
def add_entries():                       #to append all data and add entries on click the button
    a=" "
    f=firstname.get()
    f1=f.lower()
    l=lastname.get()
    l1=l.lower()
    d=dept.get()
    d1=d.lower()
    de=designation.get()
    de1=de.lower()
    list1=list(a)
    list1.append(f1)
    list1.append(l1)
    list1.append(d1)
    list1.append(de1)
    emp_dict(list1)


def add_info():                                           #for taking user input to add the enteries
    frame2.pack_forget()
    frame3.pack_forget()
    emp_first_name=Label(frame1,text="Enter first name of the employee: ",bg="red",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e1=Entry(frame1,textvariable=firstname)
    e1.grid(row=1,column=2,padx=10)
    e1.focus()
    emp_last_name=Label(frame1,text="Enter last name of the employee: ",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e2=Entry(frame1,textvariable=lastname)
    e2.grid(row=2,column=2,padx=10)
    emp_dept=Label(frame1,text="Select department of employee: ",bg="red",fg="white")
    emp_dept.grid(row=3,column=1,padx=10)
    dept.set("Select Option")
    e4=OptionMenu(frame1,dept,"Select Option","IT","Operations","Sales")
    e4.grid(row=3,column=2,padx=10)
    emp_desig=Label(frame1,text="Select designation of Employee: ",bg="red",fg="white")
    emp_desig.grid(row=4,column=1,padx=10)
    designation.set("Select Option")
    e5=OptionMenu(frame1,designation,"Select Option","Manager","Asst Manager","Project Manager","Team Lead","Senior Tester", 
                  "Junior Tester","Senior Developer","Junior Developer","Intern")
    e5.grid(row=4,column=2,padx=10)
    button4=Button(frame1,text="Add entries",command=add_entries)
    button4.grid(row=5,column=2,pady=10)
    
    frame1.configure(background="Red")
    frame1.pack(pady=10)
    
def clear_all():             #for clearing the entry widgets
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()

    
def remove_emp():                #for taking user input to remove enteries
    clear_all()
    emp_first_name=Label(frame2,text="Enter first name of the employee",bg="red",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e6=Entry(frame2,textvariable=remove_firstname)
    e6.grid(row=1,column=2,padx=10)
    e6.focus()
    emp_last_name=Label(frame2,text="Enter last name of the employee",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e7=Entry(frame2,textvariable=remove_lastname)
    e7.grid(row=2,column=2,padx=10)
    remove_button=Button(frame2,text="Click to remove",command=remove_entry)
    remove_button.grid(row=3,column=2,pady=10)
    frame2.configure(background="Red")
    frame2.pack(pady=10)

def remove_entry():  #to remove entry from excel sheet
    rsf=remove_firstname.get()
    rsf1=rsf.lower()
    print(rsf1)
    rsl=remove_lastname.get()
    rsl1=rsl.lower()
    print(rsl1)
    workbook_name="sample.xlsx"
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==rsf1 and row_value[2]==rsl1):
            print(row_value)
            print("found")
            file="sample.xlsx"
            x=pd.ExcelFile(file)
            dfs=x.parse(x.sheet_names[0])
            dfs=dfs[dfs['First Name']!=rsf]
            dfs.to_excel("sample.xlsx",sheet_name='Employee',index=False)
            messagebox.showinfo("Done","Successfully removed the Employee record")
    clear_all()

def search_emp():     #can implement search by 1st name,last name,emp id, designation
    clear_all()
    emp_first_name=Label(frame3,text="Enter first name of the employee",bg="red",fg="white")   #to take user input to seach
    emp_first_name.grid(row=1,column=1,padx=10)
    e8=Entry(frame3,textvariable=searchfirstname)
    e8.grid(row=1,column=2,padx=10)
    e8.focus()
    emp_last_name=Label(frame3,text="Enter last name of the employee",bg="red",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e9=Entry(frame3,textvariable=searchlastname)
    e9.grid(row=2,column=2,padx=10)
    search_button=Button(frame3,text="Click to search",command=search_entry)
    search_button.grid(row=3,column=2,pady=10)
    
    frame3.configure(background="Red")
    frame3.pack(pady=10)

    
def search_entry():
    sf=searchfirstname.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=searchlastname.get()
    ssl1=sl.lower()
    print(ssl1)
    path="sample.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==ssf1 and row_value[2]==ssl1):
            print(row_value)
            print("found")
            messagebox.showinfo("Done","Searched Employee Exist")
            clear_all()
    #else:
    if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
        print("Not found")
        messagebox.showerror("Sorry","Employee Record does not Exist")
        clear_all()

        
#Main window buttons and labels
def dashboard(): 
    loginFrame.pack_forget()       
    label1=Label(root,text="EMPLOYEE RECORD SYSTEM")
    label1.config(font=('Italic',16,'bold'), justify=CENTER, background="Orange",fg="Yellow", anchor="center")
    label1.pack(fill=X,padx=30, pady=30)

    label2=Label(f,text="Select an action: ",font=('bold',12), background="Black", fg="White")
    label2.pack(side=LEFT,pady=10)
    button1=Button(f,text="Add", background="Brown", fg="White", command=add_info, width=8)
    button1.pack(side=LEFT,ipadx=20,pady=10)
    button2=Button(f,text="Remove", background="Brown", fg="white", command=remove_emp, width=8)
    button2.pack(side=LEFT,ipadx=20,pady=10)
    button3=Button(f,text="Search", background="Brown", fg="White", command=search_emp, width=8)
    button3.pack(side=LEFT,ipadx=20,pady=10)
    button6=Button(f,text="Close", background="Brown", fg="White", width=8, command=root.destroy)
    button6.pack(side=LEFT,ipadx=20,pady=10)
    f.configure(background="Black")
    f.pack()

def delete_password_not_recognised():
    password_not_recog_screen.destroy()

def login_verification():

    

    username1 = username_verify.get()
    password1 = password_verify.get()
    # this will delete the entry after login button is pressed
    # username_login_entry.delete(0, END)
    # password_login_entry.delete(0, END)

    if (username1=="admin") and (password1=="admin1234") :
        dashboard()
    else:
        global password_not_recog_screen
        password_not_recog_screen = Toplevel(root)
        password_not_recog_screen.title("Error")
        password_not_recog_screen.geometry("250x250")
        Label(password_not_recog_screen, text="Invalid Credentials ").pack()
        Button(password_not_recog_screen, text="OK", command=delete_password_not_recognised).pack()
        







loginFrame.pack(pady=200,padx=20)

label1=Label(loginFrame,text="EMPLOYEE RECORD SYSTEM")
label1.config(font=('Italic',16,'bold'), background="white", justify=CENTER,fg="BLUE", anchor="center")
label1.pack(pady=10,padx=10)
# root = Toplevel(f)
# root.title("Login")
# root.geometry("300x250")

l1=Label(loginFrame, text="Please enter details below to login")
l1.config( background="white")
l1.pack()

l2=Label(loginFrame, text="")
l2.config( background="white")
l2.pack()


username_verify = StringVar()
password_verify = StringVar()


l3=Label(loginFrame, text="Username * ")
l3.config( background="white")
l3.pack()


username_login_entry = Entry(loginFrame, textvariable=username_verify)
username_login_entry.pack()

l4=Label(loginFrame, text="")
l4.config( background="white")
l4.pack()




l5=Label(loginFrame, text="Password * ")
l5.config( background="white")
l5.pack()

password__login_entry = Entry(loginFrame, textvariable=password_verify, show= '*')
password__login_entry.pack()

l6=Label(loginFrame, text="")
l6.config( background="white")
l6.pack()


button1=Button(loginFrame, text="Login", width=10, height=1, command=login_verification).pack(pady=10)
button2=Button(loginFrame,text="Close", background="Brown", fg="White", width=10, height=1, command=root.destroy).pack(pady=10)






root.mainloop()
