import math

import cv2 as cv
import matplotlib as plt
import numpy as np
import openpyxl


def rdimg(src):
    img = cv.imread(src)
    cv.imshow("img", img)
    cv.waitKey(0)
    cv.destroyAllWindows()


if __name__ == "__main__":
    src = "./photos/sample1.jpg"
    rdimg(src)
