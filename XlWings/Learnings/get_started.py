# connection with workbook
import xlwings as xw

wb = xw.Book() # open new workbook
wb = xw.Book('File.xlsx') # connect a file that is open or in the current directory
wb = xw.Book(r'path\file.xlsx') # on Windows: use raw strings to escape backslashes

# If you have the same file open in two instances of Excel, 
# you need to fully qualify it and include the app instance. 
# You will find your app instance key (the PID) via xw.apps.keys():

xw.apps[10559].books['filename.xlsx']

# Instantiate a sheet object:
sheet = wb.sheets['Sheet1']

# Reading/writing values to/from ranges is as easy as:
sheet['A1'].value = "Foo 1"
print(sheet['A1'].value)

# There are many Convenience features available e.g Range Expanding:
sheet['A1'].value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
sheet['A1'].expand().value

# Powerful converters handle most data types of interst, 
# including Numpy arrays and Pandas DataFrames in both directions:
import pandas as pd
df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
sheet['A1'].value = df
sheet['A1'].options(pd.DataFrame, expand='table').value

# Matplolib figures can be shown as pictures in Excel:
import matplotlib.pyplot as plt
fig = plt.figure()
plt.plot([1,2,3,4,5])
sheet.pictures.add(fig, name='MyPlot', update=True)
