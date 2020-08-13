#!/usr/bin/python3
# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from tkinter import *

root = Tk()
root.title("Меню")
name = StringVar()
name2 = StringVar()
name_label2 = Label(text="Введите название файла (Пример - АСУ.xlsx) :")
name_label2.grid(row=0, column=0, sticky="w")
name_entry2 = Entry(textvariable=name2,width=80)
name_entry2.grid(row=2, column=1, padx=5, pady=10)


name_label = Label(text="Введите путь до расположения паспортов (Пример C:/Users/ИД_РЦ :")
name_label.grid(row=2, column=0, sticky="w")
name_entry = Entry(textvariable=name,width=80)
name_entry.grid(row=0, column=1, padx=5, pady=5)
message_button = Button(text="Click Me")
message_button.grid(row=3,column=1, padx=5, pady=5, sticky="e")

root.mainloop()


asutp = name2.get()
doc_path = name.get

filename = asutp

print(filename)
#book = load_workbook(filename)


