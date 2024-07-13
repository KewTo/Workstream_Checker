import pygetwindow as gw
import pyautogui as auto
import time
import pandas as pd
from itertools import chain
import sys


# Grab application Workstream
def workstream():
    win = gw.getWindowsWithTitle('EXTRA! X-treme')[0]
    win.activate()


#  Read a list of desired lots from an Excel file
def read_text():
    filename = r'C:\Users\kevinto\OneDrive - Intel Corporation\Desktop\Python Dispatch.xlsx'
    df = pd.read_excel(filename, index_col=None, usecols='A', header=None)
    df = df.values.tolist()
    df = list(chain(*df))
    return df


# Run the application automatically for ease of reading
def run_workstream():
    auto.press('num1')
    auto.press('num1')
    auto.typewrite('LTHL')
    auto.press('enter')
    for LOTS in read_text():
        time.sleep(1)
        auto.typewrite(LOTS)
        auto.press('enter')
        for remaining in range(20, 0, -1):
            sys.stdout.write("\r")
            sys.stdout.write("{:2d} seconds remaining.".format(remaining))
            sys.stdout.flush()
            time.sleep(1)
        workstream()
        time.sleep(1)
        auto.press('num1')
        auto.press('tab')


# Set the application back after running the script
def back_to_default():
    for index in range(5):
        auto.press('num1')
    time.sleep(1)
    auto.press('alt')
    time.sleep(0.5)
    auto.press('T')
    time.sleep(0.5)
    auto.press('1')
    time.sleep(2)


def main()
    workstream()
    time.sleep(1)
    run_workstream()
    time.sleep(1)
    workstream()
    time.sleep(1)
    back_to_default()


if __name__ == '__main__':
    main()
