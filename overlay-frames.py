import copy
import ctypes
import datetime
import math
import os
import pathlib
import sys
import time
import tkinter
import tkinter.filedialog
import tkinter.messagebox

import cv2 as cv
import matplotlib as plt
import numpy as np
import openpyxl as opxl
import pyautogui


def showmsgbox(title, content):
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showinfo(title, content)


def choosefiles():
    root = tkinter.Tk()
    root.withdraw()
    fType = [("", "*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo("overlay-frames.py", "処理する動画ファイルを選択")
    files = tkinter.filedialog.askopenfilenames(
        filetypes=fType, initialdir=iDir
    )

    if not files:
        tkinter.messagebox.showinfo(
            "[Error] overlay-frames.py", "動画ファイルが選択されませんでした"
        )
        quit()

    files = list(files)
    # print(files)

    return files


def getcapinfo(cap):
    width, height, fps = (
        int(cap.get(cv.CAP_PROP_FRAME_WIDTH)),
        int(cap.get(cv.CAP_PROP_FRAME_HEIGHT)),
        int(cap.get(cv.CAP_PROP_FPS)),
    )

    return width, height, fps


def resizecap(cap):
    if cap.isOpened() == False:
        print("Error in opening video stream or file")
        showmsgbox("[Error] cv2", "Error in opening video stream or file")

    width, height = (
        int(cap.get(cv.CAP_PROP_FRAME_WIDTH)),
        int(cap.get(cv.CAP_PROP_FRAME_HEIGHT)),
    )

    displaysize_x, displaysize_y = pyautogui.size()
    enbale_displaysize_x, enable_displaysize_y = (
        displaysize_x * 0.8,
        displaysize_y * 0.8,
    )

    if (width > enbale_displaysize_x) or (height > enable_displaysize_y):
        if width >= height:
            display_width = enbale_displaysize_x
            resize_scale = enbale_displaysize_x / width
            display_height = round(height * resize_scale)
        else:
            display_height = enable_displaysize_y
            resize_scale = enable_displaysize_y / height
            display_width = round(width * resize_scale)
    else:
        display_width = width
        display_height = height
        resize_scale = 1.0

    return (int(display_width), int(display_height), resize_scale)


def readvideo(src, filename):
    window_title = "overlay-frames.py [{}]".format(src)

    cap = cv.VideoCapture(src)

    if cap.isOpened() == False:
        print("Error in opening file")
        showmsgbox("[Error] cv2", "Error in opening file")

    width, height, fps = getcapinfo(cap)
    display_width, display_height, resize_scale = resizecap(cap)

    overlayed_frame = np.zeros((height, width, 3), np.uint8)
    frame_avg = None

    frame_number = 0

    while cap.isOpened():
        ret, frame = cap.read()
        if not ret:
            break

        current_frame = copy.copy(frame)
        current_frame_gray = cv.cvtColor(frame, cv.COLOR_BGR2GRAY)

        if frame_avg is None:
            frame_avg = current_frame_gray.copy().astype("float")
            continue

        cv.accumulateWeighted(current_frame_gray, frame_avg, 0.6)
        frame_diff = cv.absdiff(
            current_frame_gray, cv.convertScaleAbs(frame_avg)
        )

        current_frame_resized = cv.resize(
            current_frame, dsize=(display_width, display_height)
        )
        cv.imshow(
            "{} *processing...".format(window_title), current_frame_resized
        )

        tmp = cv.addWeighted(
            src1=overlayed_frame,
            alpha=(frame_number / (frame_number + 1)),
            src2=current_frame,
            beta=(1 / (frame_number + 1)),
            gamma=0,
        )
        overlayed_frame = tmp

        overlayed_frame_resized = cv.resize(
            overlayed_frame, dsize=(display_width, display_height)
        )
        cv.imshow(window_title, overlayed_frame_resized)

        frame_number += 1

        # Press esc to exit
        if cv.waitKey(1) & 0xFF == 27:
            break

    cap.release()
    cv.destroyAllWindows()

    return


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass

    srcs = choosefiles()
    for src in srcs:
        filename = pathlib.Path(src).stem
        readvideo(src, filename)
