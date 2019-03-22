#Excel operations moduls
import openpyxl

#Gui programming moduls
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import messagebox

#Time moduls
import time
import datetime
from dateutil import parser

#Other
from re import split
import random

database = openpyxl.load_workbook('Receptek.xlsx')
#database.sheetnames
sheet = database['Sheet1']

def save_recipe():
    global all_recipe

    now = time.localtime(time.time())
     
    for row in sheet['A']:
        row_by_row = str(row.coordinate[1:])
        if sheet['A' + row_by_row].value == 2:
            sheet['A' + row_by_row].value = 1
            break
                                     
    sheet['B' + str(row_by_row)].value = e1.get()
    sheet['C' + str(row_by_row)].value = e2.get()
    sheet['D' + str(row_by_row)].value = time.strftime("%Y-%m-%d", now)

    e1.delete(first=0, last=20)
    e2.delete(first=0, last=20)
                             
    database.save("C:\JanosBuczko\Python\scipts\Cook\Receptek.xlsx")

    messagebox.showinfo("Információ", 'Sikeresen hozzáadtad a receptet a listádhoz')

    all_recipe +=1

def search():
    Recipe_name = []
    Recipe_link = []
    Recipe_date = []

    for row in sheet['B']:
        row_by_row = str(row.coordinate[1:])
        Recipe_name.append(sheet['B' + row_by_row].value)
        if sheet['A' + row_by_row].value == 2:
            break

    for row in sheet['C']:
        row_by_row = str(row.coordinate[1:])
        Recipe_link.append(sheet['C' + row_by_row].value)
        if sheet['A' + row_by_row].value == 2:
            break

    for row in sheet['D']:
        row_by_row = str(row.coordinate[1:])
        Recipe_date.append(sheet['D' + row_by_row].value)
        if sheet['A' + row_by_row].value == 2:
            break
    
    print(Recipe_name[1:-1])
    print(Recipe_link[1:-1])
    print(Recipe_date[1:-1])

    print(type(Recipe_date[0]))

    print(random.choice(Recipe_name[1:-1]))
    print(all_recipe)

def sum_recipe():
    global all_recipe
    for row in sheet['A']:
        row_by_row = str(row.coordinate[1:])
        if sheet['A' + row_by_row].value == 2:
            all_recipe = int(row_by_row)-2
            return



#GUI---------------------------------

if __name__ == "__main__":

    all_recipe = 0

    sum_recipe()

    master = tk.Tk()
    master.title('Mit főzzek ma?')
    master.geometry('500x500')

    e1 = Entry(master)
    e1.grid(row=1, column=1)
    e2 = Entry(master)
    e2.grid(row=2, column=1)

    Label(master, text ="Mentett receptjeid száma: {}".format(all_recipe)).grid(row=0, sticky=W)
    Label(master, text="Recept nev:").grid(row=1, sticky=W)
    Label(master, text="Recept link:").grid(row=2, sticky=W)
    Label(master, text="Mit főzzek?").grid(row=4, sticky=W)

    Button(master, text='Mentés', command=save_recipe).grid(row=3, column=1)
    Button(master, text='Mondd meg', command=search).grid(row=5, column=0)

    master.mainloop()

#---------------------------------


