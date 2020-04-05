from tkinter import *
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter.messagebox import showinfo
from tkinter import font
import xlsxwriter


#  Can change the name of the excel sheet here
workbook = xlsxwriter.Workbook('Journal-Entries.xlsx')  # Creates Excel File
worksheet = workbook.add_worksheet()  # Creates Work sheet

accountValue1 = []  # Debit Entry Values
accountValue2 = []  # Credit Entry Values
accountNames = []  # Debit Entry Account Names
accountNames2 = []  # Credit Entry Account Names

# Hold labels objects in array
labelList = []
labelList2 = []
labelList3 = []
labelList4 = []
labelList5 = []
yearList = []
monthList = []
dayList = []


def postTransactions():
    if len(accountNames)==0:
        showinfo("Warning", "Please Enter Atleast Once transaction Before Posting")
    else:
        showinfo("Notification", "Journal Entries have been posted Excel")
        Ledger = Tk()
        Ledger.title("Transactions")
        Ledger.geometry("300x400")
        i = 0
        master.destroy()  # Destroy the first window
        createExcel()


        if len(accountNames) > i:
            while len(accountNames) > i:

                tempValue1 = ttk.Label(Ledger, text=accountNames[i])
                labelList.append(tempValue1)

                tempValue2 = ttk.Label(Ledger,text=accountValue1[i])
                labelList2.append(tempValue2)

                tempValue3 = ttk.Label(Ledger, text=accountNames2[i])
                labelList3.append(tempValue3)

                tempValue4 = ttk.Label(Ledger, text=accountValue2[i])
                labelList4.append(tempValue4)

                tempValue5 = ttk.Label(Ledger, text=  monthList[i] + "  " + dayList[i] + " " + yearList[i])
                labelList5.append(tempValue5)
                i = i + 1

            e = 1
            y = 0
            z = len(accountValue2) + len(accountValue1)

            while z > e:
                labelList[y].grid(row=e)
                labelList5[y].grid(row=e, column=3)
                labelList2[y].grid(row=e, column=1)
                labelList3[y].grid(row=(e+1))
                labelList4[y].grid(row=(e+1), column=1)
                e = e + 2
                y = y + 1

            accountLabel = Label(Ledger,text="Account Names",bg='blue', fg='white', underline=5).grid(row=0, column=0)
            dateLabel = Label(Ledger, text="Transaction Dates", bg='blue', fg='white', underline=5).grid(row=0, column=3)

            Ledger.grid_columnconfigure(0, weight=1)
            Ledger.grid_columnconfigure(1, weight=1)
            Ledger.grid_columnconfigure(2, weight=1)
            Ledger.grid_columnconfigure(3, weight=1)
            Ledger.grid_columnconfigure(4, weight=1)
            Ledger.grid_columnconfigure(5, weight=1)

            Ledger.grid_rowconfigure(0, weight=1)
            Ledger.grid_rowconfigure(1, weight=1)
            Ledger.grid_rowconfigure(2, weight=1)
            Ledger.grid_rowconfigure(3, weight=1)
            Ledger.grid_rowconfigure(4, weight=1)
            Ledger.grid_rowconfigure(5, weight=1)
            Ledger.mainloop() # Opens second and final window


def createExcel ():

    index = 0
    placement = 2
    worksheet.write('A1', 'Account Names')
    worksheet.write('B1', 'Values')
    worksheet.write('C1', 'Transaction Date')
    while index < len(accountNames):
        column = str(placement)
        nextColumn = str(placement + 1)
        worksheet.write('A' + column, accountNames[index])
        worksheet.write('A' + nextColumn, accountNames2[index])

        worksheet.write('B' + column, accountValue1[index])
        worksheet.write('B' + nextColumn, accountValue2[index])

        worksheet.write('C' + column, monthList[index] + "  " + dayList[index] + " " + yearList[index])
        index += 1
        placement += 2

    workbook.close()


# Checks All User Inputs For Mistakes
def checkValues ():


    testValue1 = e1.get()  #
    testValue2 = e2.get()
    debitName = variable.get()
    creditName = variable2.get()
    yearName = year.get()
    dateName = date.get()
    monthName = month.get()

    if debitName == "Debit Account" or creditName == "Credit Account" or yearName == "Year" or monthName == "Month" or dateName == "Date":
        showinfo("Warning", "Please Select A Value For All Fields")
    else:
        if testValue1 == "" or testValue2 == "":
            showinfo("Warning", "There is a blank value")
        else:
            if testValue1 != testValue2:
                showinfo("Warning", "Values Are Not Equal!")
            else:
                numberCrunch()


def numberCrunch():

    temporaryValue1 = e1.get()
    temporaryValue2 = e2.get()
    temporaryValue3 = variable.get()
    temporaryValue4 = variable2.get()

    yearDate = year.get()
    monthDate = month.get()
    dayDate = date.get()

    accountValue1.append(temporaryValue1)
    accountValue2.append(temporaryValue2)  # Stores account names and balances in seperate lists,organized by index starting at 0
    accountNames.append(temporaryValue3)
    accountNames2.append(temporaryValue4)
    yearList.append(yearDate)
    monthList.append(monthDate)
    dayList.append(dayDate)

    e1.delete(0, tk.END)
    e2.delete(0, tk.END)


master = Tk()
master.title("AutoKeeper 2020")
master.geometry("500x150")


variable = StringVar(master)
variable.set("Debit")  # default value

variable2 = StringVar(master)
variable2.set("Credit")

month = StringVar(master)
month.set("Month")

year = StringVar(master)
year.set("Year")

date = StringVar(master)
date.set("Day")

master.grid_columnconfigure(0, weight=1)
master.grid_columnconfigure(1, weight=1)
master.grid_columnconfigure(2, weight=1)
master.grid_columnconfigure(3, weight=1)
master.grid_columnconfigure(4, weight=1)
master.grid_columnconfigure(5, weight=1)

master.grid_rowconfigure(0, weight=1)
master.grid_rowconfigure(1, weight=1)
master.grid_rowconfigure(2, weight=1)
master.grid_rowconfigure(3, weight=1)
master.grid_rowconfigure(4, weight=1)
master.grid_rowconfigure(5, weight=1)

debitAccounts = ttk.OptionMenu(master, variable, "Debit Account", "Inventory", "Prepaid Expense", "Accounts Receivable",
                               "Cash", "Property, Plant & Equiptment", "Notes Receivable", "Accounts Payable", "Accrued Liability",
                                "Revenue", "Cost of Goods Sold", "Unearned Revenue").grid(row=0)

creditAccounts = ttk.OptionMenu(master, variable2, "Credit Account", "Accounts Payable", "Accrued Liability",
                                "Revenue", "Cost of Goods Sold", "Unearned Revenue","Inventory", "Prepaid Expense",
                                "Accounts Receivable", "Cash", "Property, Plant & Equiptment", "Notes Receivable").grid(row=1)


yearMenu = ttk.OptionMenu(master, year, "Year", "2020", "2019", "2018").grid(row=0, column=3)
monthMenu = ttk.OptionMenu(master, month,"Month", "Jan", "Feb",
                           "Mar", "Apr", "May", "Jun", "July", "Aug",
                           "Sept", "Oct", "Nov", "Dec").grid(row=0, column=4)


dateMenu = ttk.OptionMenu(master, date, "Date", "1", "2","3", "4", "5", "6",
                                        "7", "8", "9", "10", "11", "12",
                                        "13", "14","15","16", "17", "18",
                                        "19", "20", "21", "22", "23", "24",
                                        "25", "26", "27", "28", "29", "30", "31").grid(row=0, column=5)

e1 = ttk.Entry(master)
e2 = ttk.Entry(master)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)

ttk.Button(master, text="Finish", command=postTransactions).grid(row= 5, column=0, pady=4)
ttk.Button(master, text="Next Entry", command=checkValues).grid(row=5, column=1, pady=4)
master.mainloop()

