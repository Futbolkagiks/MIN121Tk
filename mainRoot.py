from tkinter import *
import tkinter
from typing import *
import csv as v
from openpyxl import *
from colorama import init, Fore, Back, Style
import pandas as pd
from tkinter import messagebox
from auth_defs import *

root = Tk()
root.title("MIN-1-21 Project")
root.iconbitmap("image.ico")
root.withdraw()

workbook=load_workbook(filename="Users.xlsx")
EmployeesSheet=workbook["Employees"]
ClientsSheet=workbook["Clients"]
VVX=load_workbook(filename="Tariffs.xlsx")
Tariffs=VVX["Tariffs"]

Reg_Auth(root)

root.mainloop()
