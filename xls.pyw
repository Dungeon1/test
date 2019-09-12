"""Soft by HackTheSystem"""
from tkinter import *
import tkinter.scrolledtext as st
from tkinter import filedialog
import xlrd
import matplotlib.pyplot as plt
import numpy as np
import logging

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
    global mas
    mas.clear()
    comp.clear()
    sheet = rb.sheet_by_index(5)
    for i in range(3,7):
        for j in range(6,17):
            cell = sheet.cell(i,j)
            if cell.value != '':
                comp.append(cell.value)
    st.configure(state='normal')
    st.insert(END,"[Logs] Файл "+file+" добавлен\n")
    st.configure(state='disabled')    
   
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
    st.configure(state='normal')
    st.insert(END,"[Logs] Добавлен стандартный диапазон\n")
    st.configure(state='disabled')         
    graph()

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
    st.configure(state='normal')
    st.insert(END,"[Logs] Добавлен выбранный диапазон: "+st1+":"+st2+"\n")
    st.configure(state='disabled')                   
          
def graph():
    h =[]
    d = dict.fromkeys(comp)
    for i in comp:
        d[i]= mas.count(i)
        h.append(mas.count(i))
    x = np.arange(len(comp))
    plt.bar(x, height=h)
    plt.xticks(x, comp, rotation=90)
    st.configure(state='normal')
    st.insert(END,"[Logs] График построен\n")
    st.configure(state='disabled') 
    plt.show()

def clear():
    global mas
    mas.clear()
    st.configure(state='normal')
    st.insert(END,"[Logs] Все данные очищены\n")
    st.configure(state='disabled')    

root = Tk()
root.title('Учебный план')

f_top = Frame()
b1 = Button(f_top, text="Выбрать файл", width=12, height=2, command=fileopen).pack(side=LEFT)
b2 = Button(f_top, text="Построить", width=10, height=2, command=graph).pack(side=LEFT)
b3 = Button(f_top, text="Добавить диапазон", width=15, height=2, command=data).pack(side=LEFT)
b4 = Button(f_top, text="Построить автоматически", width=21, height=2, command=auto).pack(side=LEFT)
b5 = Button(f_top, text="Очистить", width=8, height=2, command=clear).pack(side=LEFT)
f_top.pack()

f_main = Frame()

f3 = LabelFrame(f_main, text="Номер листа")
t3 = Entry(f3, width=6, font='Arial 13')
t3.pack(fill=X)
f3.pack(padx=5, side=LEFT)

f1 = LabelFrame(f_main, text="Левая верхняя ячейка")
t1 = Entry(f1, width=6, font='Arial 13')
t1.pack(fill=X)
f1.pack(padx=5, side=LEFT)

f2 = LabelFrame(f_main, text="Правая нижняя ячейка")
t2 = Entry(f2, width=6, font='Arial 13')
t2.pack(fill=X)
f2.pack(padx=5, side=LEFT)

f_main.pack()

f4 = Frame()
st = st.ScrolledText(f4, width=60, height=10, font='Arial 12', state = 'disabled')
st.pack()
f4.pack(pady=4)

root.mainloop()
