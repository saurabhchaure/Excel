import xlwings as xw
import numpy as np

def world():
    wb = xw.Book.caller()
    wb.sheets[0]['A1'].value = "Hello World!"