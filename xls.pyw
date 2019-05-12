import xlrd
import matplotlib.pyplot as plt
import numpy as np
rb = xlrd.open_workbook('План.xls',formatting_info=True)
sheet = rb.sheet_by_index(5)
mas=[]
for i in range(11,572):
    for j in range(6,17):
        cell = sheet.cell(i,j)
        if cell.value != '':
            mas.append(cell.value)
comp = []           
for i in range(3,7):
    for j in range(6,17):
        cell = sheet.cell(i,j)
        if cell.value != '':
            comp.append(cell.value)
h =[]
d = dict.fromkeys(comp)
for i in comp:
    d[i]= mas.count(i)
    h.append(mas.count(i))

x = np.arange(len(comp))
plt.bar(x, height=h)
plt.xticks(x, comp, rotation=90)
plt.show()