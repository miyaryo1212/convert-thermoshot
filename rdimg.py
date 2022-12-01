import math
import datetime

import cv2 as cv
import matplotlib as plt
import numpy as np
import openpyxl as opxl


def rdimg(src):
    img = cv.imread(src)
    img_hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    h, s, v = cv.split(img_hsv)

    cv.imshow("img", img)
    #cv.imshow("img hsv", img_hsv)
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
    src = "./photos/sample1.jpg"
    v = rdimg(src)
    path = "./output/saved_{}.xlsx".format(
        datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    )
    cvtlist(v, path)
