from os import read
from matplotlib import pyplot as plt
from pandas.core.algorithms import mode
from typing import Tuple
import csv
from openpyxl import *
import pandas as pd
from colorama import init, Fore, Back, Style
from tkinter import *
from tkinter import messagebox
import numpy as np


workbook=load_workbook(filename="Users.xlsx")
EmployeesSheet=workbook["Employees"]
ClientsSheet=workbook["Clients"]
VVX=load_workbook(filename="Tariffs.xlsx")
Tariffs=VVX["Tariffs"]

def Close(the_window):
    the_window.destroy()
    return

def showTariffs(window):
    count=0
    for yes in Tariffs.iter_rows(values_only=True):
        ID=Label(window,text=yes[0],bd=2,bg="grey").grid(column=0,row=count)
        TariffName=Label(window,text=yes[1],bd=2,bg="grey").grid(column=1,row=count)
        Data=Label(window,text=yes[2],bd=2,bg="grey").grid(column=2,row=count)
        Time=Label(window,text=yes[3],bd=2,bg="grey").grid(column=3,row=count)
        Price=Label(window,text=yes[4],bd=2,bg="grey").grid(column=4,row=count)
        count+=1
    return

def showMyTariff(window,details):
    ActiveT=""
    PreviousT=""
    print(details)
    for col in Tariffs.iter_rows(min_col=1,max_col=3,min_row=2,values_only=True):
        if col[0]==details[5]:
            ActiveT=col[1]
    for col in Tariffs.iter_rows(min_col=1,max_col=3,min_row=2,values_only=True):
        if col[0]==details[6]:
            PreviousT=col[1]
    Lbl1=Label(window,text=f"Your active Tariff is {ActiveT}").grid(column=0,row=0)
    Lbl2=Label(window,text=f"Your previous Tariff was {PreviousT}").grid(column=1,row=0)
    CloseButton=Button(window,text="Close",command=lambda w=window: Close(w)).grid(column=0,row=1)
    
def users_list(window,type):
    count=0
    for yes in type.iter_rows(values_only=True):
        ID=Label(window,text=yes[0],bd=2,bg="grey").grid(column=0,row=count)
        NAME=Label(window,text=yes[1],bd=2,bg="grey").grid(column=1,row=count)
        LOGIN=Label(window,text=yes[2],bd=2,bg="grey").grid(column=2,row=count)
        EXTRA=Label(window,text=yes[4],bd=2,bg="grey").grid(column=3,row=count)
        count+=1
    CloseButton=Button(window,text="Close",command=lambda w=window: Close(w)).grid()
    return

def searchClient(window,Search):
    name=Search.get()
    found=False
    while True:
        for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
            if (name.upper() in col[1].upper())==True:
                found=True
                ID=Label(window,text=col[0],bd=2,bg="grey").grid(column=0,row=2)
                NAME=Label(window,text=col[1],bd=2,bg="grey").grid(column=1,row=2)
                LOGIN=Label(window,text=col[2],bd=2,bg="grey").grid(column=2,row=2)
                BALANCE=Label(window,text=col[4],bd=2,bg="grey").grid(column=3,row=2)
                if col[7]!=None and col[8]!=None:
                    AGE=Label(window,text=col[7],bd=2,bg="grey").grid(column=4,row=2)
                    CITY=Label(window,text=col[8],bd=2,bg="grey").grid(column=5,row=2)
        if found==False:
            messagebox.showinfo("Error", "Client could not be found")
            continue
        elif found==True:
            break

def searchClientWindow(window,Search):
    SearchLabel=Label(window,text="Enter the name of a Client").grid(column=0,row=0)
    SearchField=Entry(window,textvariable=Search).grid(column=1,row=0)
    SearchButton=Button(window,text="Enter",command=lambda w=window,s=Search: searchClient(w,s)).grid(column=0,row=1)

def historyUser(window,Search):
    ID=Search.get()
    for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
        if ID==int(col[0]):
            Lbl1=Label(window,text=f"{col[1]}'s active Tariff is {col[5]}").grid(column=0,row=2)
            Lbl2=Label(window,text=f"{col[1]}'s previous Tariff was {col[6]}").grid(column=0,row=3)

def historyUserWindow(window,Search):
    SearchLabel=Label(window,text="Enter the ID of a Client").grid(column=0,row=0)
    SearchField=Entry(window,textvariable=Search).grid(column=1,row=0)
    SearchButton=Button(window,text="Enter",command=lambda w=window,s=Search: historyUser(w,s)).grid(column=0,row=1)
    
def sortClients():
    print(
    'ID - 1\n'
    'Name - 2\n'
    'Age - 3\n'
    'City - 4\n')
    c = input("Enter by what you wish to sort clients: ")
    b = pd.read_excel('Users.xlsx',sheet_name="Clients")
    lst = []
    for i in range(len(b["Id"])):
        lst.append([b["Id"][i], b["Name"][i], b["Age"][i], b["City"][i]])
    if c == '1':
        lst = sorted(lst, key=lambda x: x[0])
        for i in lst:
            print(i)
    elif c == '2':
        lst = sorted(lst, key=lambda x: x[1])
        for i in lst:
            print(i)
    elif c == '3':
        lst = sorted(lst, key=lambda x: x[2])
        for i in lst:
            print(i)
        if input('statik or graphik') == 'stitik':
            lst = sorted(lst, key=lambda x: x[2])
            for i in lst:
                print(i)
        else:
            gp = pd.read_excel('Users.xlsx')
            old = 0
            jun = 0
            med = 0
            for i in range(len(gp['Age'])):
                if gp['Age'][i] < 18:
                    jun += 1
                elif 18 <= gp['Age'][i] < 50:
                    med += 1
                else:
                    old += 1
            data = [
                ['Young', jun],
                ['Mature', med],
                ['Old', old],
            ]
            values = [x[1] for x in data]
            labels = [x[0] for x in data]
            fig, ax = plt.subplots()
            fig.canvas.set_window_title('Inai')
            ax.set_title("MIN-1-21")

            ax.pie(values, labels=labels, autopct="%.1f%%", radius=1.2)

            ax.set_aspect("equal")

            plt.show()
    elif c== "4":
        lst = sorted(lst, key=lambda x: x[3])
        for i in lst:
            print(i)
    return

def stats(window):
    age = []
    local = []
    b = pd.read_excel('Users.xlsx',sheet_name="Clients")
    for i in range(len(b['Age'])):
        # age.append(b['Age'][i])
        local.append(b['City'][i])
    # age1 = set(age)
    local1 = set(local)
    lbl=Label(window,text="Stats by Region").pack()
    for i in local1:
        q1=local.count(i), i
        lbl1=Label(window,text=q1).pack()
    # lbl2=Label(window,text="Stats by Age").pack()
    # for i in age1:
    #     q2=age.count(i), i
    #     lbl3=Label(window,text=q1).pack()

def addBalance(Search,Amount,w):
    s=Search.get()
    a=Amount.get()
    count=1
    for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
        count+=1
        if s==col[0]:
            ClientsSheet[f"E{count}"]=int(col[4])+a
            messagebox.showinfo("Balance","Money has been deposited")
            workbook.save("Users.xlsx")
    return

def addBalanceWindow(window,Search,Amount):
    SearchLabel=Label(window,text="Enter the ID of a Client").grid(column=0,row=0)
    SearchField=Entry(window,textvariable=Search).grid(column=1,row=0)
    AmountLabel=Label(window,text="Enter the Amount").grid(column=0,row=1)
    AmountField=Entry(window,textvariable=Amount).grid(column=1,row=1)
    SearchButton=Button(window,text="Enter",command=lambda s=Search,a=Amount,w=window: addBalance(s,a,w)).grid(column=0,row=2)
    CloseButton=Button(window,text="Close",command=lambda w=window: Close(w)).grid(column=0,row=3)
    
def subscribeToNewTariff(window,details,idTariff):
    thefile=pd.DataFrame([details[0],idTariff])
    thefile.to_csv("Applications.csv",mode='a',header=False)
    Response=Label(window,text="Your request has been submitted").grid(column=0,row=2)

def subscribeToNewTariffWindow(window,idTariff,details):
    TariffLabel=Label(window,text="Enter the ID of a Tariff:").grid(column=0,row=0)
    TariffField=Entry(window,textvariable=idTariff).grid(column=1,row=0)
    TariffButton=Button(window,text="Enter", command=lambda w=window, d=details,id=idTariff: subscribeToNewTariff(w,d,id)).grid(column=0,row=1)
    CloseButton=Button(window,text="Close",command=lambda w=window: Close(w)).grid(column=0,row=1)

def RequestSubFunction1(id):
    for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
        if id==col[0]:
            return col[1]

def RequestSubFunction2(id):
    for row in Tariffs.iter_rows(min_row=2,values_only=True):
        if row[0]==id:
            return row[1]

def viewListOfReqest(window):
    thefile=pd.read_csv("Applications.csv")
    count=0
    for row in range(len(thefile['clientId'])):
        q=[(thefile['clientId'][row]),(thefile['tariffId'][row])]
        print(q)
        infoClient=RequestSubFunction1(q[0])
        infoTariff=RequestSubFunction2(q[1])
        RequestLabel=Label(window,text=f"User {infoClient} wishes to use Tariff {infoTariff}").grid(column=0,row=count)
        ApproveButton=Button(window,text="Approve",command=lambda t="A",qq=q,c=count: analysisRequest(t,qq,c)).grid(column=1,row=count)
        RejectButton=Button(window,text="Reject",command=lambda t="R",qq=q,c=count: analysisRequest(t,qq,c)).grid(column=2,row=count)
        count+=1
    CloseButton=Button(window,text="Close",command=lambda w=window: Close(w)).grid(column=0,row=1)
    
def analysisRequest(t,qq,c):
    workbook=load_workbook(filename="Users.xlsx")
    ClientsSheet=workbook["Clients"]
    thefile=pd.read_csv("Applications.csv")
    if t=="A":
        count=1
        for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
            count+=1
            if qq[0]==col[0]:
                ClientsSheet[f"G{count}"]=col[5]
                ClientsSheet[f"F{count}"]=qq[1]
                workbook.save("Users.xlsx")
    messagebox.showinfo("Request","Request has been processed")
    thefile.drop(c)
    thefile.to_csv("Applications.csv")

def addInfoToClient(details):
    print("Extra INFO screen")
    add_info=[input("Enter the city where you live: "), input("Enter your age: ")]
    count=1
    for col in ClientsSheet.iter_rows(min_row=2,values_only=True):
        count+=1
        if int(details[0])==int(col[0]):
            ClientsSheet[f"H{count}"]=add_info[1]
            ClientsSheet[f"I{count}"]=add_info[0]
            ClientsSheet[f"E{count}"]=0
            workbook.save("Users.xlsx")
    print("Extra informations has been saved")

def graf():
    gs = pd.read_excel('Users.xlsx')
    n = [0, 0, 0, 0, 0, 0, 0, ]
    for i in range(len(gs['City'])):
        if gs['City'][i] == 'Osh':
            n[0] += 1
        elif gs['City'][i] == 'Batken':
            n[1] += 1
        elif gs['City'][i] == 'Jalal-Abad':
            n[2] += 1
        elif gs['City'][i] == 'Chuy':
            n[3] += 1
        elif gs['City'][i] == 'Issuk-Kul':
            n[4] += 1
        elif gs['City'][i] == 'Talas':
            n[5] += 1
        elif gs['City'][i] == 'Naryn':
            n[6] += 1
        index = np.arange(7)

    plt.title('A Bar Chart')
    plt.bar(index, n, error_kw={'ecolor': '0.1', 'capsize': 8}, alpha=0.9, label='Regions')
    plt.xticks(index, ['Osh', 'Batken', 'Jalal-Abad', 'Chuy', 'Issuk-Kul', 'Talas', 'Naryn'])
    plt.legend(loc=2)
    plt.show()

def createEmployee(n,l,p):
    workbook=load_workbook(filename="Users.xlsx")
    EmployeesSheet=workbook["Employees"]
    account_type=EmployeesSheet
    new_account=[int(len(account_type["A"])),n.get(),l.get(),p.get()]
    account_type.append(new_account)
    workbook.save("Users.xlsx")

def createEmployeeWindow(window,Login,Password,Name):
    NameField=Entry(window, textvariable=Name).pack()
    LoginField=Entry(window, textvariable=Login).pack()
    PasswordField=Entry(window, textvariable=Password).pack()
    Label1=Label(window, textvariable="Name").pack(side=NameField.LEFT)
    Enter_Button = Button(window, text="Enter",command=lambda t="Clients", n=Name,l=Login,p=Password: createEmployee(n,l,p)).pack()