# from tkinter import *
# import tkinter
# from typing import *
# import csv as v
# from openpyxl import *
# from colorama import init, Fore, Back, Style
# import pandas as pd
# from tkinter import messagebox
# Name=StringVar()
# Login=StringVar()
# Password=StringVar()

# def create_user(type,window):
#     workbook=load_workbook(filename="Users.xlsx")
#     EmployeesSheet=workbook["Employees"]
#     ClientsSheet=workbook["Clients"]
#     if type=="Clients":
#         account_type=ClientsSheet
#     elif type=="Employees":
#         account_type=EmployeesSheet
#     new_account=[int(len(account_type["A"])),Name.get(),Login.get(),Password.get()]
#     if login_check_2(account_type,window)==False:
#         account_type.append(new_account)
#     workbook.save("Users.xlsx")
    
# def login_check_2(type,window):
#     LoginCheck=False
#     while True:
#         auth1=Login.get()
#         for col in type.iter_rows(min_row=2,values_only=True):
#             if auth1==col[2]:
#                 LoginCheck=True
#         if LoginCheck==True:
#             messagebox.showinfo("Error", "This Login is already in use!")
#         else:
#             messagebox.showinfo("Success","New account has been created")
#             window.destroy()
#         break
#     return LoginCheck