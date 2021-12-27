from tkinter import *
import tkinter
from typing import *
import csv as v
from openpyxl import *
from colorama import init, Fore, Back, Style
import pandas as pd
from tkinter import messagebox
from menus import *
# from menu_functions import *

workbook=load_workbook(filename="Users.xlsx")
EmployeesSheet=workbook["Employees"]
ClientsSheet=workbook["Clients"]
VVX=load_workbook(filename="Tariffs.xlsx")
Tariffs=VVX["Tariffs"]
root = Tk()
root.withdraw()

Name=StringVar()
Login=StringVar()
Password=StringVar()

def create_user(type,window):
    workbook=load_workbook(filename="Users.xlsx")
    EmployeesSheet=workbook["Employees"]
    ClientsSheet=workbook["Clients"]
    if type=="Clients":
        account_type=ClientsSheet
    elif type=="Employees":
        account_type=EmployeesSheet
    new_account=[int(len(account_type["A"])),Name.get(),Login.get(),Password.get()]
    if login_check_2(account_type,window)==False:
        account_type.append(new_account)
    workbook.save("Users.xlsx")
    
def login_check_2(type,window):
    LoginCheck=False
    while True:
        auth1=Login.get()
        for col in type.iter_rows(min_row=2,values_only=True):
            if auth1==col[2]:
                LoginCheck=True
        if LoginCheck==True:
            messagebox.showinfo("Error", "This Login is already in use!")
        else:
            messagebox.showinfo("Success","New account has been created")
            window.destroy()
        break
    return LoginCheck

def check(t,top,top2,a):
    workbook=load_workbook(filename="Users.xlsx")
    EmployeesSheet=workbook["Employees"]
    ClientsSheet=workbook["Clients"]
    auth1=Login.get()
    auth2=Password.get()
    CHECK1=False
    CHECK2=False
    if t=="Client":
        account_type=ClientsSheet
    elif t=="Employee":
        account_type=EmployeesSheet
    for col in account_type.iter_rows(min_row=2,values_only=True):
        if col[2]==auth1:
            CHECK1=True
        if col[3]==auth2:
            CHECK2=True
            global details
            details=col
    if (CHECK1==True) and (CHECK2==True):
        print("Success")
        top.destroy()
        top2.destroy()
        a.destroy()
        if t=="Client":
            users_menu(details)
        elif t=="Employee":
            if auth1=="Dir123" and auth2=="Ctor321":
                dirs_menu(details)
            else:
                emps_menu(details)
    else:
        print("Failure")

def account_auth(type,ttt,a):
    enterwindow = Toplevel()
    enterwindow.title("MIN-1-21 Project")
    enterwindow.iconbitmap("image.ico")
    enterwindow.geometry("300x300")
    LoginField=Entry(enterwindow, textvariable=Login).pack()
    PasswordField=Entry(enterwindow, textvariable=Password, show='*').pack()
    Enter_Button = Button(enterwindow, text="Enter",command=lambda t=type,tt=enterwindow,t4=ttt,aa=a: check(t,tt,t4,aa)).pack()

def type_choice(a):
    authwindow = Toplevel()
    authwindow.geometry("300x300")
    authwindow.title("MIN-1-21 Project")
    authwindow.iconbitmap("image.ico")
    Emp_Button = Button(authwindow, text="Employee", command=lambda ttt=authwindow, m="Employee",aa=a: account_auth(m,ttt,aa))
    Emp_Button.pack()
    Cli_Button = Button(authwindow, text="Client", command=lambda ttt=authwindow, m="Client",aa=a: account_auth(m,ttt,aa))
    Cli_Button.pack()

def reg_client(a):
    enterwindow = Toplevel()
    enterwindow.title("MIN-1-21 Project")
    enterwindow.iconbitmap("image.ico")
    enterwindow.geometry("300x300")
    NameField=Entry(enterwindow, textvariable=Name).pack()
    LoginField=Entry(enterwindow, textvariable=Login).pack()
    PasswordField=Entry(enterwindow, textvariable=Password).pack()
    Label1=Label(enterwindow, textvariable="Name").pack(side=NameField.LEFT)
    Enter_Button = Button(enterwindow, text="Enter",command=lambda t="Clients", w=enterwindow: create_user(t,w)).pack()

def Reg_Auth(root):
    RegAuthwindow=Toplevel()
    RegAuthwindow.geometry("300x300")
    RegAuthwindow.title("MIN-1-21 Project")
    RegAuthwindow.iconbitmap("image.ico")
    Reg_Button = Button(RegAuthwindow, text="Registration", command=lambda a=RegAuthwindow: reg_client(a))
    Reg_Button.pack()
    Auth_Button = Button(RegAuthwindow, text="Authentication", command=lambda a=RegAuthwindow: type_choice(a))
    Auth_Button.pack()

