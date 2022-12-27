import ctypes
import datetime
import os
import pathlib
import pprint
import time
import tkinter
import tkinter.filedialog
import tkinter.messagebox
from tkinter import ttk

import openpyxl as opxl
from openpyxl.utils import get_column_letter


def showmsgbox(title, content):
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showinfo(title, content)


def choosefile():
    root = tkinter.Tk()
    root.withdraw()
    fType = [("", ".xlsx")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo("compare-sheets", "Excelファイル (.xlsx形式) を選択")
    file = tkinter.filedialog.askopenfilename(filetypes=fType, initialdir=iDir)

    if not file:
        tkinter.messagebox.showinfo(
            "[Error] compare-sheets", "Excelファイル (.xlsx形式) が選択されませんでした"
        )
        quit()

    return file


def detectsheets(booksrc):
    workbook = opxl.load_workbook(booksrc)
    sheets = workbook.sheetnames

    return(sheets)


def choosesheets(list):
    root = tkinter.Tk()
    root.withdraw()

    root.title("choose-sheets")
    root.geometry("300x130")

    def ok_get(event):
        root.quit()

    label_before = tkinter.Label(root, text="Before")
    label_before.grid(column=0, row=0, padx=20, pady=10)
    label_after = tkinter.Label(root, text="After")
    label_after.grid(column=0, row=1, padx=20, pady=10)

    combo_before = ttk.Combobox(root, values=list, justify="center")
    combo_before.grid(column=1, row=0, padx=20, pady=10)
    combo_before.set(list[0])
    combo_after = ttk.Combobox(root, values=list, justify="center")
    combo_after.grid(column=1, row=1, padx=20, pady=10)
    combo_after.set(list[1])


    button = tkinter.Button(root, text="OK")
    button.bind("<Button-1>", ok_get)
    button.grid(column=1, row=2, padx=20, pady=10)

    root.bind("<Return>", ok_get)

    root.deiconify()
    root.mainloop()
    root.withdraw()

    return combo_before.get(), combo_after.get()


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass

    src = choosefile()
    sheets = detectsheets(src)
    print(choosesheets(sheets))
