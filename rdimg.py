import datetime
import math
import os
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
    tkinter.messagebox.showinfo("rdimg.py", "F30サーモショットで撮影した320x240の画像を選択")
    file = tkinter.filedialog.askopenfilename(filetypes=fType, initialdir=iDir)

    if not file:
        tkinter.messagebox.showinfo("rdimg.py", "画像を選択してください。プログラムを終了します。")
        quit()

    return file


def rdimg(src, min, max):
    img = cv.imread(src)
    img_hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    h, s, v = cv.split(img_hsv)

    for y in range(240):
        for x in range(320):
            v[y][x] = min + ((max - min) / 255 * v[y][x])

    cv.imshow("img", img)
    # cv.imshow("img hsv", img_hsv)
    cv.waitKey(0)
    cv.destroyAllWindows()

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
    src = choosefile()

    v = rdimg(src, 20, 100)

    time_now = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    path = "./output/saved_{}.xlsx".format(time_now)
    cvtlist(v, path)

    print('Saved in to "saved_{}.xlsx"'.format(time_now))
    tkinter.messagebox.showinfo(
        "rdimg.py", 'Saved in to "saved_{}.xlsx"'.format(time_now)
    )
