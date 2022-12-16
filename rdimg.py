import ctypes
import datetime
import math
import os
import pathlib
import sys
import tkinter
import tkinter.filedialog
import tkinter.messagebox

import cv2 as cv
import matplotlib as plt
import numpy as np
import openpyxl as opxl


def choosefile():
    root = tkinter.Tk()
    root.withdraw()
    fType = [("", "*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo("rdimg.py", "F30サーモショットで撮影した画像を選択")
    files = tkinter.filedialog.askopenfilenames(
        filetypes=fType, initialdir=iDir
    )

    if not files:
        tkinter.messagebox.showinfo("[Error] rdimg.py", "画像が選択されませんでした")
        quit()

    files = list(files)
    # print(files)

    return files


def askminmax():
    root = tkinter.Tk()
    root.withdraw()

    root.title("設定温度を小数第1位まで入力")
    root.geometry("400x120")

    def ok_get(event):
        root.quit()

    label_min = tkinter.Label(root, text="最小温度：")
    label_max = tkinter.Label(root, text="最大温度：")
    label_min.place(x=20, y=20)
    label_max.place(x=20, y=50)

    editbox_min = tkinter.Entry(root, width=40)
    editbox_min.insert(tkinter.END, "")
    editbox_min.place(x=100, y=20)

    editbox_max = tkinter.Entry(root, width=40)
    editbox_max.insert(tkinter.END, "")
    editbox_max.place(x=100, y=50)

    button = tkinter.Button(root, text="OK")
    button.bind("<Button-1>", ok_get)
    button.place(x=300, y=80)

    root.bind("<Return>", ok_get)

    root.deiconify()
    root.mainloop()
    root.withdraw()

    if not (editbox_min.get() or editbox_max.get()):
        tkinter.messagebox.showinfo("rdimg.py", "設定温度が正しく入力されませんでした")
        quit()

    min, max = float(editbox_min.get()), float(editbox_max.get())

    return min, max


def rdimg(src, min, max):
    print("[info] {} loading".format(src))
    img = cv.imread(src)
    img_hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    h, s, v = cv.split(img_hsv)

    print("[info] {} processing".format(src))
    for y in range(len(v)):
        for x in range(len(v[0])):
            v[y][x] = min + ((max - min) / 255 * v[y][x])

    # cv.imshow("rdimg.py - Press Esc to continue", img)
    # cv.imshow("img hsv", img_hsv)
    # cv.waitKey(0)
    # cv.destroyAllWindows()
    print("[info] {} finished".format(src))

    return v


def cvtlist(list, path):
    wb = opxl.Workbook()
    sheet = wb.worksheets[0]

    for y in range(240):
        for x in range(320):
            list_value = list[y][x]
            sheet.cell(y + 1, x + 1, list_value)

    """
    for row_num in range(240):
        sheet.row_dimensions[row_num + 1].height = 20

    for column_num in range(320):
        sheet.column_dimensions[column_num + 1].width = 20
    """

    wb.save(path)


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass

    srcs = choosefile()
    temp_min, temp_max = askminmax()

    saved_filenames = []

    for src in srcs:
        v = rdimg(src, temp_min, temp_max)

        dir = "./output"
        if not os.path.exists(dir):
            os.makedirs(dir)

        src = pathlib.Path(src)

        # time_now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        path = "./output/file_{}.xlsx".format(src.stem)
        cvtlist(v, path)

        saved_filenames.append("file_{}.xlsx".format(src.stem))

    save_message = "Saved into:"
    for name in saved_filenames:
        save_message += "\n{}".format(name)

    tkinter.messagebox.showinfo("rdimg.py", save_message)
