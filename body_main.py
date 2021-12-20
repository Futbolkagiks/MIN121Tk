from tkinter import *
from typing import *
import csv as v
from openpyxl import *
from colorama import init, Fore, Back, Style
import pandas as pd

root = Tk()

workbook=load_workbook(filename="Users.xlsx")
EmployeesSheet=workbook["Employees"]
ClientsSheet=workbook["Clients"]
VVX=load_workbook(filename="Tariffs.xlsx")
Tariffs=VVX["Tariffs"]

Emp_Button = Button(root, text="Employee").pack()
Cli_Button = Button(root, text="Client").pack()

root.mainloop()

