'''
sheet=book.active
sheet['A1'] = 56
sheet['A2'] = 43
now = date.today()
sheet['A3'] = now
    
sheet = book.active
rows = [
        [88, 46, 90],
        [89, 38, 12],
        [23, 59, 78],
        [56, 0, 98],
        [24, 0, 43],
        [34, 0, 67]
    ]
rows = ["1","2","3"]
sheet.append(rows)
book.save('expenses.xlsx')

book=load_workbook("expenses.xlsx")
sheet=book.active
a=sheet["A1"]
print(sheet["A1"].value)

'''
from openpyxl import *
from openpyxl.styles import *
from datetime import date
from time import sleep
import os

def createTable():
    '''Creates an excel file and types formatted titles'''
    book = Workbook()
    sheet=book.active
    sheet["A1"],sheet["B1"],sheet["C1"],sheet["D1"]=\
            "Num","Purpose","Cost","Date"
    for cell in ["A1","B1","C1","D1"]:
        sheet[cell].font=Font(bold=True)
        sheet[cell].alignment=Alignment(horizontal="center",vertical="center")
    book.save("expenses.xlsx")
    return book

def appendTable(lst):##
    sheet = book.active
    sheet['A1'] = "fe3"
    sheet.cell(row=3, column=2).value = "fdf"
    book.save('expenses.xlsx')

def checkDate(date):
    d,m,y=int(date[0]),int(date[1]),int(date[2])
    if date(y,m,d)>date.today():
        return False
    if date[1] in ["01","03","05","07","08","10","12"]:
        if int(date[0])>31:
            return False
    elif date[1] in ["04","06","09","11"]:
        if int(date[0])>30:
            return False
    else:
        if int(date[2])%4==0:
            if int(date[0])>29:
                return False
        else:
            if int(date[0])>28:
                return False
    return True


    

def main():
    book=createTable()
    info=[]
    while True:
        print("LIST OF EXPENCES")
        purpose=input("Type a purpose/reason for expenditure:").strip()
        cost=input("Type a cost of it:(in dollars)")
        while True:
            date=input("Type date of it:(e.g DD.MM.YYYY)").split(".")#it is a list
            if len(date[0])!=2 or len(date[1])!=2 or len(date[2])!=4:
                print("The given date is incorrect!")
            else:
                break
            
        info.append([purpose,cost,date])
        check=input("Do you want to add something?:(split line if NO!)")
        sleep(1.2)
        os.system("cvd")
        if len(check)==0:
            break
    


if __name__=="__main__":
    main()
