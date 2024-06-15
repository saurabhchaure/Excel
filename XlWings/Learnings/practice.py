import xlwings as xw
import pandas as pd


wb = xw.Book(r'F:\Excel\XlWings\Learnings\practice.xlsx')
sheet = wb.sheets[0]

# df = pd.DataFrame([[1,2], [3,4]], columns=['a', 'b'])
# sheet['A1'].value = df
# print(sheet['A1'].options(pd.DataFrame, expand='table').value)
# wb.save()

import matplotlib.pyplot as plt

fig = plt.figure()
plt.plot([1,2,3,4,5])
sheet.pictures.add(fig, name='MyPlot', update=True)

