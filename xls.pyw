from tkinter import *
from tkinter import filedialog
import xlrd
import matplotlib.pyplot as plt
import numpy as np

rb = xlrd.book.Book()
dictt = {'A':1,'B':2,'C':3,'D':4,'E':5,'F':6,'G':7,'H':8,
    'I':9,'J':10,'K':11,'L':12,'M':13,'N':14,'O':15,'P':16,'Q':17,
    'R':18,'S':19,'T':20,'U':21,'V':22,'W':23,'X':24,'Y':25,'Z':26
    }
mas = []
comp = []

def fileopen():
    file = filedialog.askopenfilename(initialdir="/Рабочий стол",title="Select file",filetypes=(("Excel","*.xls"),("all files","*.*")))
    global rb
    rb = xlrd.open_workbook(file, formatting_info=True)
    global comp
    sheet = rb.sheet_by_index(5)
    for i in range(3,7):
        for j in range(6,17):
            cell = sheet.cell(i,j)
            if cell.value != '':
                comp.append(cell.value) 
   
def auto():
    global mas
    global comp
    sheet = rb.sheet_by_index(5)
    mas=[]
    for i in range(11,572):
        for j in range(6,17):
            cell = sheet.cell(i,j)
            if cell.value != '':
                mas.append(cell.value)
    for i in range(588,605):
        for j in range(6,17):
            cell = sheet.cell(i,j)
            if cell.value != '':
                mas.append(cell.value)           
    for i in range(630,638):
        for j in range(6,17):
            cell = sheet.cell(i,j)
            if cell.value != '':
                mas.append(cell.value)           

def data():
    global mas
    st1 = t1.get()
    st2 = t2.get()
    ind = int(t3.get())-1
    column1 = dictt[st1[0]]-1
    row1 = int(st1[1:])-1
    column2 = dictt[st2[0]]-1
    row2 = int(st2[1:]) 
    sheet = rb.sheet_by_index(ind)    
    for i in range(row1, row2):
        for j in range(column1,column2):
            cell = sheet.cell(i,j)
            if cell.value != '':
                mas.append(cell.value)                
          
def graph():
    h =[]
    d = dict.fromkeys(comp)
    for i in comp:
        d[i]= mas.count(i)
        h.append(mas.count(i))

    x = np.arange(len(comp))
    plt.bar(x, height=h)
    plt.xticks(x, comp, rotation=90)
    plt.show()


root = Tk()
root.title('Учебный план')

f_top = Frame()
b1 = Button(f_top, text="Выбрать файл", width=15, height=2, command=fileopen).pack(side=LEFT)
b2 = Button(f_top, text="Построить", width=15, height=2, command=graph).pack(side=LEFT)
b3 = Button(f_top, text="Добавить диапазон", width=15, height=2, command=data).pack(side=LEFT)
b4 = Button(f_top, text="Автоматический диапазон", width=15, height=2, command=auto).pack(side=LEFT)
f_top.pack()

f_data = Frame()
L1 = Label(f_data, text="Левая верхняя ячейка").pack(side=TOP, pady=2)
t1 = Entry(f_data, width=6, font='Arial 13')
t1.pack()
L2 = Label(f_data, text="Правая нижняя ячейка").pack(side=TOP, pady=2)
t2 = Entry(f_data, width=6, font='Arial 13')
t2.pack()
L1 = Label(f_data, text="Номер листа").pack(side=TOP, pady=2)
t3 = Entry(f_data, width=6, font='Arial 13')
t3.pack()
f_data.pack()

root.mainloop()
