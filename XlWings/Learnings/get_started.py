#connection with workbook
import xlwings as xw

wb = xw.Book() # open new workbook
wb = xw.Book('File.xlsx') # connect a file that is open or in the current directory
wb = xw.Book(r'path\file.xlsx') # on Windows: use raw strings to escape backslashes

# If you have the same file open in two instances of Excel, 
# you need to fully qualify it and include the app instance. 
# You will find your app instance key (the PID) via xw.apps.keys():

