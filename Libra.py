#-------------------------------------------------------------------------------
# Name:        module2
# Purpose:
#
# Author:      black
#
# Created:     19/06/2019
# Copyright:   (c) black 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import json
import re
import logging
from pathlib import Path
from urllib.request import urlopen, Request
try:
	from mailmerge import *
except:
	pass
import getpass
try:
    import datetime
    from threading import Thread
    import time
    from tkinter import *
    from tkinter import ttk
    from tkinter.filedialog import askopenfilename
    from tkinter.filedialog import asksaveasfilename
    from tkinter.filedialog import askdirectory
    import os
    import sqlite3
    import tkinter.messagebox
    import xlrd
except:
    pass

conn = sqlite3.connect('books.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()
#create tables if they dont exit
conn.execute("create table if not exists book(book_id   INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,book_title TEXT NOT NULL,book_author TEXT NOT NULL,book_no TEXT NOT NULL,level TEXT NOT NULL DEFAULT 'general',subject TEXT NOT NULL DEFAULT 'general',class TEXT NOT NULL DEFAULT 'general')")
conn.execute("create table if not exists borrow(id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,file_id INTEGER,st_name TEXT,st_class TEXT,stream TEXT,dob TEXT,dor	TEXT DEFAULT 'None')")

#header for view books
view_books = ('TITLE','AUTHOR','BOOK NO')
#student borrowed book
view_detail = ('NAME','BOOK No','DATE TAKEN','DATE RETURNED')

user = getpass.getuser()

import_st = ('NAME','ClASS','STREAM','BOOK_No','DOB')

path = "c:\\users\\"+user+"\\documents\\"
file = path+"books.xlsx"
file2 = path+"student.xlsx"

try:

	book = open('books.xlsx','rb')
	student = open('student.xlsx','rb')
	booktemp = open(file,'wb')
	studenttemp = open(file2,'wb')

	for m in student:
		studenttemp.write(m)
	studenttemp.close()
	student.close()

	for x in book:
		booktemp.write(x)
	booktemp.close()
	book.close()
except:
	pass


class sms:
    def __init__(self,master):
        master.title("Libra Mini")
        master.geometry("{}x{}".format(master.winfo_screenwidth() - 100, master.winfo_screenheight() - 100))
        #master.geometry("750x500+0+0")
        master.iconbitmap("accessories-dictionary-2.ico")
        self.rows = 0
        while self.rows < 50:
            master.rowconfigure(self.rows,weight=1)
            master.columnconfigure(self.rows,weight=1)
            self.rows +=1
    #menu
    #style

        menu = Menu(master)
        master.config(menu=menu)
        newMenu = Menu(menu)
        menu.add_cascade(label="New",menu=newMenu)
        newMenu.add_command(label="Add Book          Control+A",command=self.Add_book_Gui)
        newMenu.add_command(label="Borrow               Control+B",command=self.borrow_book_Gui)
        newMenu.add_command(label="Student Clr        Control+C",command=self.clear_st_Gui)
        newMenu.add_separator()
        newMenu.add_command(label="Exit")
    #edit menu
        #editMenu = Menu(menu)
        #menu.add_cascade(label="Edit",menu=editMenu)
        #editMenu.add_command(label="Delete book",command=self.delete_book)
        #editMenu.add_command(label="Delete Student",command=self.delete_student)# command=self.search)

    #Report
        reportMenu = Menu(menu)
        menu.add_cascade(label="Reports",menu=reportMenu)
        reportMenu.add_command(label="Process Reports",command=self.generate)
        reportMenu.add_separator()
        
    #Help menu
        helpMenu = Menu(menu)
        menu.add_cascade(label="Help",menu=helpMenu)
        helpMenu.add_command(label="document",command=self.help_Gui)
        helpMenu.add_command(label='Contact',command=self.contact)
        helpMenu.add_separator()
        helpMenu.add_command(label="About Libra",command=self.about)
    #notebook
        self.notebook = ttk.Notebook(master)
        self.notebook.grid(row=0,column=0, columnspan=50, rowspan=50, sticky='NESW')

        self.dashboard = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard, text=' DashBoard ')

        self.books = ttk.Frame(self.notebook)
        self.notebook.add(self.books,text=' Libra Books  ')

        self.others = ttk.Frame(self.notebook)
        self.notebook.add(self.others, text='  Libra List  ')

        self.importBks = ttk.Frame(self.notebook)
        self.notebook.add(self.importBks, text=' Import Books ')

        self.importSts = ttk.Frame(self.notebook)
        self.notebook.add(self.importSts, text=' Import Students ')

        #self.style = ttk.style()
        #self.style.config("mystyle.Treeview", font=('Calibri', 11))
        #self.style.config("mystyle.Treeview.Heading",font=('Calibri',13,'bold'))


    #import books
        self.row = 0
        while self.row < 50:
            self.importBks.rowconfigure(self.row,weight=1)
            self.importBks.columnconfigure(self.row, weight=1)
            self.row +=1

        self.sheet = StringVar()
        self.sheet.set('sheet1')
        self.brFrame = Frame(self.importBks,bg='skyblue')

        self.brExcel = ttk.Button(self.brFrame,text='Browse Excel',command=self.bulk_books)
        self.brExcel.grid(row=1,column=1,pady=2,padx=2,sticky='')
        self.brFile = ttk.Entry(self.brFrame, width=50)
        self.brFile.grid(row=1,column=50,pady=2,padx=2,sticky='NESW')
        self.clear = ttk.Button(self.brFrame,text='Clear',command=self.cleartv)
        self.clear.grid(row=1,column=100,pady=2,sticky='')
        

        self.bload = ttk.Button(self.brFrame,text='Load',command=self.load_book)
        self.bload.grid(row=2,column=1,pady=2,padx=2,sticky='')
        self.exsheet = ttk.Entry(self.brFrame,textvariable=self.sheet, width=50)
        self.exsheet.grid(row=2,column=50,pady=2,padx=2,sticky='NESW')
        self.save = ttk.Button(self.brFrame,text='save',command=self.bulk_save_book)
        self.save.grid(row=2,column=100,pady=2,sticky='')
        self.brFrame.grid(row=0,column=0,columnspan=50,rowspan=2,sticky='NEWS')


        container = Frame(self.importBks)
        container.grid(row=2,column=0,columnspan=50,rowspan=48,sticky='NEWS')

        self.imp_view = ttk.Treeview(columns=view_books, show="headings")
        imp_vsb = ttk.Scrollbar(orient="vertical", command=self.imp_view.yview)
        imp_hsb = ttk.Scrollbar(orient="horizontal", command=self.imp_view.xview)
        self.imp_view.configure(yscrollcommand=imp_vsb.set, xscrollcommand=imp_hsb.set)
        self.imp_view.configure(height=15)
        self.imp_view.grid(column=0, row=0, sticky='nsew', in_=container)
        imp_vsb.grid(column=1, row=0, sticky='ns', in_=container)
        imp_hsb.grid(column=0, row=1, sticky='ew', in_=container)
        self.imp_view.bind("<ButtonRelease-1>", lambda event, t=self.imp_view: self.select_book(t))


        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        for col in view_books:
            self.imp_view.heading(col, text=col.title())
            self.imp_view.column(col)




    #admin Dashboard
        self.row = 0
        while self.row < 50:
            self.dashboard.rowconfigure(self.row,weight=1)
            self.dashboard.columnconfigure(self.row, weight=1)
            self.row +=1
            #second row
        self.board = Frame(self.dashboard,bg="violet")
        self.board.grid(row=14,column=0,columnspan=50,rowspan=10,sticky='NEWS')
        self.lab = Label(self.board,text="BORROWED BOOKS: ",background='violet',foreground='white')
        self.lab.place(x=5,y=30)
        self.lab.config(font=('Courier',18,'bold'))
        self.lab_1 = Label(self.board,text="5034",background='violet',foreground='white')
        self.lab_1.place(x=300,y=30)
        self.lab_1.config(font=('Courier',22,'bold'))


            #first row
        self.board2 = Frame(self.dashboard,bg="brown")
        self.board2.grid(row=2,column=0,columnspan=50,rowspan=8,sticky='NEWS')
        self.lab2 = ttk.Label(self.board2,text="WELCOME TO LIBRA",background='brown',foreground='white')
        self.lab2.pack()
        self.lab2.config(font=('Courier',22,'bold'))
        #self.lab2_1 = Label(self.board2,text="675",background='brown',foreground='white')
        #self.lab2_1.place(x=300,y=30)
        #self.lab2_1.config(font=('Courier',22,'bold'))



            #third row
        self.board3 = Frame(self.dashboard,bg="green")
        self.board3.grid(row=26,column=0,columnspan=50,rowspan=10,sticky='NEWS')
        self.lab3 = ttk.Label(self.board3,text="RETURNED BOOKS: ",background='green',foreground='white')
        self.lab3.place(x=5,y=30)
        self.lab3.config(font=('Courier',18,'bold'))
        self.lab3_1 = Label(self.board3,text="2354",background='green',foreground='white')
        self.lab3_1.place(x=300,y=30)
        self.lab3_1.config(font=('Courier',22,'bold'))

            #fouth row
        self.board4 = Frame(self.dashboard,bg='black')
        self.board4.grid(row=38,column=0,columnspan=50,rowspan=10,sticky='NEWS')
        self.lab4 = ttk.Label(self.board4,text="BOOKS PRESENT: ",background='black',foreground='white')
        self.lab4.place(x=5,y=30)
        self.lab4.config(font=('Courier',18,'bold'))
        self.lab4_1 = Label(self.board4,text="109456",background='black',foreground='white')
        self.lab4_1.place(x=300,y=30)
        self.lab4_1.config(font=('Courier',22,'bold'))

        #books
        self.row = 0
        while self.row < 50:
            self.books.rowconfigure(self.row,weight=1)
            self.books.columnconfigure(self.row,weight=1)
            self.row +=1

            #search frame
        self.sframe = Frame(self.books,bg='skyblue')

        self.sbox = ttk.Entry(self.sframe,width=48)
        self.sbox.grid(row=1,column=21,columnspan=20,pady=2,padx=2,sticky='NESW')
        self.sbutton = ttk.Button(self.sframe,text='search',width=20,command=self.search_book)
        self.sbutton.grid(row=1,column=1,pady=2,padx=2,sticky='')



        self.slabel = Label(self.sframe,text='', bg='skyblue',foreground='green')
        self.slabel.grid(row=2,column=21,rowspan=4,pady=2,padx=2,stick='NEWS')
        self.slabel.config(font=('Courier',14,'bold'))

        self.All_bks = ttk.Button(self.sframe,text=' All Books ',width=20,command=self.display_books)
        self.All_bks.grid(row=2,column=1,sticky='E',pady=2,padx=2)


        self.sframe.grid(row=0,column=0,columnspan=50,rowspan=2, sticky='NEWS')
            #view all books
        container = Frame(self.books)
        container.grid(row=2,column=0,columnspan=50,rowspan=48,sticky='NEWS')

        self.b_view = ttk.Treeview(columns=view_books, show="headings")
        b_vsb = ttk.Scrollbar(orient="vertical", command=self.b_view.yview)
        b_hsb = ttk.Scrollbar(orient="horizontal", command=self.b_view.xview)
        self.b_view.configure(yscrollcommand=b_vsb.set, xscrollcommand=b_hsb.set)
        self.b_view.configure(height=15)
        self.b_view.grid(column=0, row=0, sticky='nsew', in_=container)
        b_vsb.grid(column=1, row=0, sticky='ns', in_=container)
        b_hsb.grid(column=0, row=1, sticky='ew', in_=container)
        self.b_view.bind("<ButtonRelease-3>", lambda event, t=self.b_view: self.select_book(t))


        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        for col in view_books:
            self.b_view.heading(col, text=col.title())
            self.b_view.column(col,width=20)
        self.display_books()
        #self.b_view.config(selectmode='browse')
        #self.b_view.bind('<<TreeviewSelect>>', self.callback(self.b_view))
        



        #student borrow data
        self.row = 0
        while self.row < 50:
            self.others.columnconfigure(self.row,weight=1)
            self.others.rowconfigure(self.row, weight=1)
            self.row +=1

            #filter search area
        self.sframe = Frame(self.others,bg='skyblue')



        self.stsearch = ttk.Button(self.sframe,text='SEARCH',width=20,command=self.student_search)
        self.stsearch.grid(row=1,column=1,padx=2,pady=2)
        self.stentry = Entry(self.sframe,width=48)
        self.stentry.grid(row=1,column=21,pady=2,padx=2,sticky='NEWS')

        self.st_all = ttk.Button(self.sframe,text='All Students',width=20,command=self.all_student)
        self.st_all.grid(row=2,column=1)


        self.slabel1 = Label(self.sframe,text='', bg='skyblue',foreground='green')
        self.slabel1.grid(row=2,column=21,rowspan=4,pady=2,padx=2,stick='NEWS')
        self.slabel1.config(font=('Courier',14,'bold'))


        self.sframe.grid(row=0,column=0,columnspan=50,rowspan=2, sticky='NEWS')
            #view all student with books
        container = Frame(self.others)
        container.grid(row=2,column=0,columnspan=50,rowspan=48,sticky='NEWS')
        self.s_view = ttk.Treeview(columns=view_detail, show="headings")
        s_vsb = ttk.Scrollbar(orient="vertical", command=self.s_view.yview)
        s_hsb = ttk.Scrollbar(orient="horizontal", command=self.s_view.xview)
        self.s_view.configure(yscrollcommand=s_vsb.set, xscrollcommand=s_hsb.set)
        self.s_view.configure(height=20)
        self.s_view.grid(column=0, row=0, sticky='nsew', in_=container)
        s_vsb.grid(column=1, row=0, sticky='ns', in_=container)
        s_hsb.grid(column=0, row=1, sticky='ew', in_=container)
        self.s_view.bind("<ButtonRelease-3>", lambda event, t=self.s_view: self.select_student(t))

        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        for col in view_detail:
            self.s_view.heading(col, text=col.title())
            self.s_view.column(col)
        self.all_student()




    #import bulk students
        self.row = 0
        while self.row < 50:
            self.importSts.rowconfigure(self.row,weight=1)
            self.importSts.columnconfigure(self.row, weight=1)
            self.row +=1

        self.sheet1 = StringVar()
        self.sheet1.set('sheet1')
        self.StFrame = Frame(self.importSts,bg='skyblue')

        self.brExcel = ttk.Button(self.StFrame,text='Browse Excel',command=self.bulk_students)
        self.brExcel.grid(row=1,column=1,pady=2,padx=2,sticky='')
        self.brFile1 = ttk.Entry(self.StFrame, width=50)
        self.brFile1.grid(row=1,column=50,pady=2,padx=2,sticky='NESW')
        self.clear = ttk.Button(self.StFrame,text='Clear')
        self.clear.grid(row=1,column=100,pady=2,sticky='')

        self.bload = ttk.Button(self.StFrame,text='Load',command = self.load_student)
        self.bload.grid(row=2,column=1,pady=2,padx=2,sticky='')
        self.exsheet1 = ttk.Entry(self.StFrame,textvariable=self.sheet, width=50)
        self.exsheet1.grid(row=2,column=50,pady=2,padx=2,sticky='NESW')
        self.save = ttk.Button(self.StFrame,text='save',command=self.bulk_save_student)
        self.save.grid(row=2,column=100,pady=2,sticky='')
        self.StFrame.grid(row=0,column=0,columnspan=50,rowspan=2,sticky='NEWS')


        container = Frame(self.importSts)
        container.grid(row=2,column=0,columnspan=50,rowspan=48,sticky='NEWS')

        self.impSt_view = ttk.Treeview(columns=import_st, show="headings")
        impSt_vsb = ttk.Scrollbar(orient="vertical", command=self.impSt_view.yview)
        impSt_hsb = ttk.Scrollbar(orient="horizontal", command=self.impSt_view.xview)
        self.impSt_view.configure(yscrollcommand=impSt_vsb.set, xscrollcommand=impSt_hsb.set)
        self.impSt_view.configure(height=15)
        self.impSt_view.grid(column=0, row=0, sticky='nsew', in_=container)
        impSt_vsb.grid(column=1, row=0, sticky='ns', in_=container)
        impSt_hsb.grid(column=0, row=1, sticky='ew', in_=container)


        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        for col in import_st:
            self.impSt_view.heading(col, text=col.title())
            self.impSt_view.column(col)



#for the dash board
    def dash(self):
        #tkinter.messagebox.showinfo("Success","Function is working perfectly")
        c.execute("select count(*) from book")
        self.result = c.fetchone()[0]
        c.execute("select count(*) from borrow where not dor='None'")
        self.result1 = c.fetchone()[0]
        #c.execute("select count(*) from borrow where dor='None'")
        #self.result2 = c.fetchone()[0]
        c.execute("select count(*) from borrow inner join book on book.book_id=borrow.file_id where dor='None'")
        self.lab_1.config(text=str(c.fetchone()[0]))
        #self.lab2_1.config(text=str(self.result2))
        self.lab3_1.config(text=str(self.result1))
        self.lab4_1.config(text=str(self.result))

#clear treeview
    def cleartv(self):
        for i in self.imp_view.get_children():
            self.imp_view.delete(i)
#search for books in book tab
    def search_book(self):
        self.search_string = self.sbox.get()

        if self.search_string=='':
             tkinter.messagebox.showwarning("Warning","Please Enter search term such as math ")
        else:
            self.search_term = self.search_string.split()
            x = 0
            self.sql = ""
            for self.term in self.search_term:
                x +=1
                if x==1:
                    self.sql = "book_title like '%"+self.term+"%' or subject like '%"+self.term+"%' or book_no like '%"+self.term+"%' or level like '%"+self.term+"%' or class like '%"+self.term+"%' or book_author like '%"+self.term+"%'"
                else:
                    self.sql +=" or book_author like '%"+self.term+"%' or book_title like '%"+self.term+"%' or subject like '%"+self.term+"%' or book_no like '%"+self.term+"%' or level like '%"+self.term+"%' or class like '%"+self.term+"%'"

            self.sql1 = "select * from book where "+self.sql
            self.sql2 = "select count(*) from book where "+self.sql

            try:
                self.all = conn.execute(self.sql1)
                c.execute(self.sql2)
                self.numb = c.fetchone()[0]
            except:
                pass

            for i in self.b_view.get_children():
                self.b_view.delete(i)
            try:
                for self.row in self.all:
                    if self.row['book_title']=='':
                         tkinter.messagebox.showinfo("Sorry","No results found for "+ self.search_string)
                    else:
                        title = self.row['book_title']
                        #subject = self.row['subject']
                        author = self.row['book_author']
                        no = self.row['book_no']
                        #clas = self.row['class']
                        values = (title,author,no)
                        self.b_view.insert("", END, values=values)
                #
                #print(self.numb)
            except:
                pass
            print(type(self.numb))
            self.slabel.config(text=str(self.numb)+" Books ")

          
#student data search
    def student_search(self):
        self.search_string = self.stentry.get()
        if self.search_string=='':
            tkinter.messagebox.showwarning("Warning","Please Enter such term like name or class")
        else:
            self.search_term = self.search_string.split()
            x=0
            self.sql = ""
            #create search query
            for l in self.search_term:
                x +=1
                if x==1:
                    self.sql = "st_name like '%"+l+"%' or st_class like '%"+l+"%' or stream like '%"+l+"%'"
                else:
                    self.sql += " or st_name like '%"+l+"%' or st_class like '%"+l+"%' or stream like '%"+l+"%'"

            self.sql1 = "select * from borrow inner join book on borrow.file_id=book.book_id where "+self.sql
            self.sql2 = "select count(*) from borrow inner join book on borrow.file_id=book.book_id where "+self.sql

            #selecting search data
            try:
                self.all = conn.execute(self.sql1)
                c.execute(self.sql2)
                num = c.fetchone()[0]
            except:
                tkinter.messagebox.showinfo("Sorry","Results not found for "+self.search_string)
            #insert search results into view
            for i in self.s_view.get_children():
                self.s_view.delete(i)

            for self.row in self.all:
                name = self.row['st_name']
                book_no = self.row['book_no']
                dob = self.row['dob']
                dor = self.row['dor']
                values = (name,book_no,dob,dor)
                self.s_view.insert("", END, values=values)
            self.slabel1.config(text=str(num)+" Results")



#display all student records
    def all_student(self):
        sql = "select * from borrow inner join book on borrow.file_id=book.book_id"
        try:
            self.all_students = conn.execute(sql)
            c.execute("select count(*) from borrow inner join book on borrow.file_id=book.book_id")
            num = c.fetchone()[0]
        except:
            pass
        for i in self.s_view.get_children():
            self.s_view.delete(i)

        for self.row in self.all_students:
            name = self.row['st_name']
            book_no = self.row['book_no']
            dob = self.row['dob']
            dor = self.row['dor']
            values = (name,book_no,dob,dor)
            self.s_view.insert("", END, values=values)
        self.slabel1.config(text=str(num)+" Students")
#call function when item is selected from treeview
    def callback(self,treeview):
        pass
#filter books in book tab
    def bulk_books(self):
        try:
            excel = askopenfilename()
            if excel.endswith('.xlsx') or excel.endswith('.xls'):
                self.brFile.delete(0,END)
                self.brFile.insert(0,excel)
            else:
                tkinter.messagebox.showwarning("Warning","Select An Excel file")
        except:
            pass
#load books into treeview
    def load_book(self):
        excel = self.brFile.get()
        if excel=='':
            tkinter.messagebox.showwarning("Warning","Please browse for excel file")
        else:
            self.sht = self.exsheet.get()
            book = xlrd.open_workbook(excel)
            try:
                sheet = book.sheet_by_name(self.sht.title())
            except:
                tkinter.messagebox.showwarning("Warning",self.sht.title()+" does not exist")
            for i in self.imp_view.get_children():
                self.imp_view.delete(i)
            try:
                for r in range(1, sheet.nrows):
                    title = sheet.cell(r,0).value
                    author = sheet.cell(r,1).value
                    book_no = sheet.cell(r,2).value
                    value = (title,author,book_no)
                    self.imp_view.insert("", END, values=value)
            except:
                pass
    def bulk_save_book(self):
        try:
            excel = self.brFile.get()
            self.sht = self.exsheet.get()
            book = xlrd.open_workbook(excel)
            sheet = book.sheet_by_name(self.sht.title())
        except:
            pass
        query = """INSERT INTO book(book_title, book_author, book_no) VALUES (?, ?, ?)"""
        if excel=='':
            tkinter.messagebox.showwarning("Warning","Please Select valid Excel file")
        else:
            try:
                k = 0
                for r in range(1, sheet.nrows):
                    k +=1
                    title = sheet.cell(r,0).value
                    author = sheet.cell(r,1).value
                    book_no = sheet.cell(r,2).value
                    value = (title,author,book_no)
                    c.execute(query, value)
                conn.commit()
                tkinter.messagebox.showinfo("Success",str(k)+" books have been added")
                self.dash()
            except:
                tkinter.messagebox.showerror("Warning","Unable to Save books")

# displaying all books in book tab
    def display_books(self):
        self.all_books = conn.execute("select * from book")
        c.execute("select count(*) from book")
        bks = c.fetchone()[0]
        for i in self.b_view.get_children():
            self.b_view.delete(i)
        try:
            for self.row in self.all_books:
                title = self.row['book_title']
                #subject = self.row['subject']
                author = self.row['book_author']
                no = self.row['book_no']
                #clas = self.row['class']
                values = (title,author,no)
                self.b_view.insert("", END, values=values)
            self.dash()
            #self.all_student()
        except:
            tkinter.messagebox.showerror("Error","Unable to fetch all the files")
        self.slabel.config(text=str(bks)+" Books")


    def bulk_students(self):
        try:
            excel = askopenfilename()
            if excel.endswith('.xlsx') or excel.endswith('.xls'):
                self.brFile1.delete(0,END)
                self.brFile1.insert(0,excel)
            else:
                tkinter.messagebox.showwarning("Warning","Select An Excel file")
        except:
            pass

    def load_student(self):
        excel = self.brFile1.get()
        if excel=='':
            tkinter.messagebox.showwarning("Warning","Please browse for excel file")
        else:
            self.sht = self.exsheet1.get()
            book = xlrd.open_workbook(excel)
            try:
                sheet = book.sheet_by_name(self.sht.title())
            except:
                tkinter.messagebox.showwarning("Warning",self.sht.title()+" does not exist")
            for i in self.impSt_view.get_children():
                self.impSt_view.delete(i)
            try:
                for r in range(1, sheet.nrows):
                    Name = sheet.cell(r,0).value
                    Clas = sheet.cell(r,1).value
                    Stream = sheet.cell(r,2).value
                    book_no = sheet.cell(r,3).value
                    dob = sheet.cell(r,4).value
                    value = (Name,Clas,Stream,book_no,dob)
                    self.impSt_view.insert("", END, values=value)
            except:
                pass

    def bulk_save_student(self):
        try:
            excel = self.brFile1.get()
            self.sht = self.exsheet1.get()
            book = xlrd.open_workbook(excel)
            sheet = book.sheet_by_name(self.sht.title())
        except:
            pass
        query = """INSERT INTO borrow(file_id, st_name, st_class, stream, dob) VALUES (?, ?, ?, ?, ?)"""
        if excel=='':
            tkinter.messagebox.showwarning("Warning","Please Select valid Excel file")
        else:
            try:
                k = 0
                for r in range(1, sheet.nrows):
                    k +=1
                    Name = sheet.cell(r,0).value
                    clas = sheet.cell(r,1).value
                    stream = sheet.cell(r,2).value
                    book_no = sheet.cell(r,3).value
                    dob = sheet.cell(r,4).value
                    c.execute("select book_id from book where book_no='"+book_no+"'")
                    result = c.fetchone()[0]
                    value = (result, Name, clas, stream, dob)
                    c.execute(query, value)
                conn.commit()
                self.dash()
                self.display_books()
                self.all_student()
                tkinter.messagebox.showinfo("Success",str(k)+" books have been added")
                self.dash()
            except:
                tkinter.messagebox.showerror("Warning","Invalid Book Nos, They dont exist in the system")


            
    def Add_book_Gui(self):
        def commit_data():
            title = Etitle.get()
            author = Eauthor.get()
            book_no = Ebook_no.get()
            if title =='' or author=='' or book_no=='':
                tkinter.messagebox.showwarning("Warning","Please fill up all the Entry fields")
                window.lift()
            else:
                value = (title,author,book_no.upper())
                query = """INSERT INTO book(book_title, book_author, book_no) VALUES (?, ?, ?)"""
                try:
                    if c.execute(query, value):
                        conn.commit()
                        tkinter.messagebox.showinfo("Success",title+" has been added")
                        self.dash()
                        self.display_books()
                        self.all_student()
                        window.destroy()
                    else:
                        tkinter.messagebox.showinfo("Sorry","Failed to add "+title)
                        window.lift()
                except:
                    pass
            #

        window = Toplevel()

        window.title('Add Book')
        window.geometry("270x200+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lTitle = Label(mainFrame,text='TITLE',width=10)
        lTitle.place(x=10,y=25)
        Etitle = Entry(mainFrame,width='25')
        Etitle.place(y=25,x=90)
        lauthor = Label(mainFrame,text='AUTHOR',width=10)
        lauthor.place(x=10,y=70)
        Eauthor = Entry(mainFrame,width='25')
        Eauthor.place(x=90,y=70)
        lBook_no = Label(mainFrame,text='BOOK No',width=10)
        lBook_no.place(x=10,y=115)
        Ebook_no = Entry(mainFrame,width=25)
        Ebook_no.place(x=90,y=115)
        Button(mainFrame,text='submit',width=18,command=commit_data).place(x=60,y=155)



        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')
        #window.iconbitmap("button-yellow.ico")
        window.mainloop()

    def borrow_book_Gui(self):
        def commit_data():
            x = datetime.datetime.now()
            x = x.strftime("%d")+"/"+x.strftime("%m")+"/"+x.strftime("%Y")
            Name = Ename.get()
            clss = clas.get()
            stream1 = stream.get()
            book_no = Ebook_no.get()
            if Name=='' or clss=='' or stream1=='' or book_no=='':
                tkinter.messagebox.showwarning("Warning","Please Fill in the required fields")
                window.lift()
            else:
                try:
                    c.execute("SELECT book_id FROM book WHERE book_no='"+book_no.upper()+"'")
                    result = c.fetchone()[0]
                except:
                    tkinter.messagebox.showinfo("Sorry",book_no+" does not exist")
                    window.lift()
                    #tkinter.messagebox.askyesno("kevin","It works")
                try:
                    value = (result, Name, clss, stream1.upper(), x)
                    conn.execute("INSERT INTO borrow (file_id, st_name, st_class, stream, dob) VALUES (?,?,?,?,?)",value)
                    conn.commit()
                    tkinter.messagebox.showinfo("Success",Name+" has borrowed file "+book_no.upper())
                    window.destroy()
                    self.dash()
                    self.all_student()
                except:
                    window.lift()


        window = Toplevel()

        window.title('Borrow')
        window.geometry("270x250+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lName = Label(mainFrame,text='NAME',width=10)
        lName.place(x=10,y=25)
        Ename = Entry(mainFrame,width='25')
        Ename.place(y=25,x=90)

        clas = StringVar()
        value = ('Teacher','S.6','S.5','S.4','S.3','S.2','S.1')
        lClass = Label(mainFrame,text='ClASS',width=10)
        lClass.place(x=10,y=70)
        Eclass = ttk.Combobox(mainFrame,textvariable=clas,values=value,width='22')
        Eclass.place(x=90,y=70)

        stream = StringVar()
        strm = ('Arts','Sciences','West','East','South','North','Central')
        lStream = Label(mainFrame,text='STREAM',width=10)
        lStream.place(x=10,y=115)

        Estream = ttk.Combobox(mainFrame,textvariable=stream,values=strm,width=22)
        Estream.place(x=90,y=115)

        lBook_no = Label(mainFrame,text='BOOK No',width=10)
        lBook_no.place(x=10,y=160)
        Ebook_no = Entry(mainFrame,width=25)
        Ebook_no.place(x=90,y=160)

        Button(mainFrame,text='submit',width=18,command=commit_data).place(x=60,y=200)



        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')

        window.mainloop()



    def clear_st_Gui(self):
        def commit_data():
            x = datetime.datetime.now()
            x = x.strftime("%d")+"/"+x.strftime("%m")+"/"+x.strftime("%Y")
            book_no = Ename.get()
            #check if entry is empty

            if book_no !='':
                try:
                    c.execute("select count(*) from book where book_no='"+book_no.upper()+"'")
                    exist = c.fetchone()[0]
                    if exist==0:
                        tkinter.messagebox.showwarning("Warning","Please Enter valid book no. "+book_no.upper()+" does not exist")
                        window.lift()
                    else:
                        c.execute("select book_id from book where book_no='"+book_no.upper()+"'")
                        result = c.fetchone()[0]
                        c.execute("select count(*) from borrow where file_id=? and dor=?",(result,'None'))
                        borrowed = c.fetchone()[0]
                        if borrowed==1:
                            sql = "Update borrow set dor='"+x+"' where file_id=? and dor=?"
                            value = (result,'None')
                            conn.execute(sql,value)
                            conn.commit()
                            name = conn.execute("select * from borrow inner join book on book.book_id=borrow.file_id where borrow.dor=? and book_no=? order by id desc limit 1",(x,book_no.upper()))
                            for row in name:
                                tkinter.messagebox.showinfo("Success",row['st_name']+" has been cleared")
                            window.destroy()
                            self.dash()
                            self.display_books()
                            self.all_student()
                        else:
                            tkinter.messagebox.showerror("Warning",book_no+" already returned, check that you entered right book no")
                            window.lift()

                except:
                    pass
            else:
                tkinter.messagebox.showinfo("Warning","Please Enter book id")
                window.lift()



        window = Toplevel()
        window.title('Clear Book')
        window.geometry("270x100+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lName = Label(mainFrame,text='BOOk No',width=10)
        lName.place(x=10,y=25)
        Ename = Entry(mainFrame,width='25')
        Ename.place(y=25,x=90)
        Button(mainFrame,text='submit',width=18,command=commit_data).place(x=60,y=55)
        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')

        window.mainloop()

       #edit student
    def delete_book(self):
        def kill_it():
            window.destroy()
        def commit_data():
            book_no = Ename.get()
            print(book_no.upper())
            try:
                book = conn.execute("select * from book where book_no='"+book_no.upper()+"'")
                for bk in book:
                    title = bk['book_title']
                    id_b = bk['book_id']
                print(title)
                print(id_b)
                conn.execute("delete from book where book_no='"+book_no.upper()+"'")
                conn.commit()
                tkinter.messagebox.showinfo("Success",title+" "+book_no.upper()+" has been deleted")
                self.dash()
                self.display_books()
                window.destroy()
            except:
                tkinter.messagebox.showinfo("Sorry",book_no.upper()+" does not exist")
                window.lift()



        window = Toplevel()
        window.title('Delete Book')
        window.geometry("270x100+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lName = Label(mainFrame,text='BOOk No',width=10)
        lName.place(x=10,y=25)
        Ename = Entry(mainFrame,width='25')
        Ename.place(y=25,x=90)
        Button(mainFrame,text='Delete',width=9,command=commit_data).place(x=60,y=55)
        Button(mainFrame,text='Cancel',width=9,command=kill_it).place(x=140,y=55)
        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')
        window.mainloop()




    #Edit book
    def delete_student(self):
        def commit_data():
            book_no = Ename.get()
            try:
            	r = conn.execute("select * from book where book_no='"+book_no+"'")
            	for x in r:
            		x= x['book_id']
            	conn.execute("delete from borrow where file_id=?",(x))
            	conn.commit()
            	tkinter.messagebox.showinfo("Success",book_no+" has been deleted")
            	window.destroy()
            except:
            	tkinter.messagebox.showinfo("Sorry",book_no+" does not exist or check your spelling")
            	window.lift()

        window = Toplevel()

        window.title('Clear Student')
        window.geometry("270x100+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lName = Label(mainFrame,text='Name',width=10)
        lName.place(x=10,y=25)
        Ename = Entry(mainFrame,width='25')
        Ename.place(y=25,x=90)
        Button(mainFrame,text='submit',width=18,command=commit_data).place(x=60,y=55)
        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')

        window.mainloop()
   #generate report
    def report_Gui(self):
        def commit_data():
            x = datetime.datetime.now()
            x = x.strftime("%d")+"/"+x.strftime("%m")+"/"+x.strftime("%Y")
            book_no = Ename.get()

        window = Toplevel()

        window.title('Clear Student')
        window.geometry("270x100+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window,bg='green')

        lName = Label(mainFrame,text='BOOk No',width=10)
        lName.place(x=10,y=25)
        Ename = Entry(mainFrame,width='25')
        Ename.place(y=25,x=90)
        Button(mainFrame,text='submit',width=18,command=commit_data).place(x=60,y=55)
        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')

        window.mainloop()

    def help_Gui(self):
        text = open("help.txt","r")
        help_txt = text.read()
        os.startfile('help.txt')






    def generate(self):
        def commit_data():
            stream1 = stream.get()
            classe1 = classes.get()
            user = getpass.getuser()
            results = ""
            kevin = []
            x = 0
            path = "c:\\users\\"+user+"\\Documents\\"
            file = path+classe1+"__"+stream1+".docx"
            if stream1!='All':
                results = conn.execute("select * from borrow inner join book on book.book_id=borrow.file_id where dor='None' and (st_class='"+classe1+"' and stream='"+stream1+"') order by st_name asc")
            elif stream1=='All' and classe1=='All':
                results = conn.execute("SELECT * FROM borrow INNER JOIN book ON book.book_id=borrow.file_id WHERE dor='None' and not st_class='Teacher' order by st_class asc")
            else:
                results = conn.execute("select * from borrow inner join book on book.book_id=borrow.file_id where st_class='"+classe1+"' and dor='None' order by st_name asc")
            for result in results:
                x +=1
                k = {'numb':str(x),'class':result['st_class'],'name':result['st_name'],'numer':result['book_no'],'title':result['book_title'],'stream':result['stream']}
                kevin.append(k)
            if(x>0):
                document = MailMerge('template.docx')
                document.merge()
                document.merge_rows('class',kevin)
                document.write(file)
                tkinter.messagebox.showinfo("Success",classe1+"__"+stream1+".docx"+" has been created in you Documents folder")
                window.destroy()
            else:
                tkinter.messagebox.showinfo("Oops!",classe1+" "+stream1+" has no Defaulters")
                window.lift()
            


        window = Toplevel()

        window.title('generate document')
        window.geometry("270x150+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        classes = StringVar()
        stream = StringVar()
        strm = ('All','Arts','Sciences','West','East','South','North','Central')
        values = ('Teacher','All','S.6','S.5','S.4','S.3','S.2','S.1')

        mainFrame = Frame(window,bg='green')

        Lclass = Label(mainFrame,text='Class',width=10)
        Lclass.place(x=10,y=25)
        Eclass = ttk.Combobox(mainFrame,textvariable=classes,values=values,width=22)
        Eclass.place(x=90,y=25)

        lStream = Label(mainFrame,text='Stream',width=10)
        lStream.place(x=10,y=60)
        Estream = ttk.Combobox(mainFrame,textvariable=stream,values=strm,width=22)
        Estream.place(x=90,y=60)

        Button(mainFrame,text='Generate',width=18,command=commit_data).place(x=60,y=90)
        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')
        window.mainloop()



    def contact(self):
        window = Toplevel()
        window.title('Contact')
        window.geometry("270x80+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)
        data = "Email: "+"babskevon@gmail.com"+"\n"+"\t"+"edrinkevin1@gmail.com"
        right = "Tel: "+"0786842944"+"/"+"0705412894"

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window)
        lab1 = Label(mainFrame,text=data)
        lab1.pack()
        #lab1.config(font=('Calibri',18,'bold'))
        lab2 = Label(mainFrame,text=right)
        lab2.pack()

        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')
        window.mainloop()




    def about(self):
        window = Toplevel()
        window.title('Delete Book')
        window.geometry("270x100+200+100")
        #window.maxsize(350,350)
        window.resizable(False,False)
        data = "Libra"
        right = "\n"+"Copyright @ Black-Tec Africa"+"\n"+"version 1.5.0"+"\n"

        row = 0
        while row < 50:
            window.rowconfigure(row,weight=1)
            window.columnconfigure(row,weight=1)
            row +=1
        mainFrame = Frame(window)
        lab1 = Label(mainFrame,text=data)
        lab1.pack()
        lab1.config(font=('Calibri',18,'bold'))
        lab2 = Label(mainFrame,text=right)
        lab2.pack()

        mainFrame.grid(row=0,column=0,rowspan=50,columnspan=50,sticky='NEWS')
        window.mainloop()

    def select_book(self,tree,event=None):
    	#print(type(tree))
    	#master.clipboard_clear()
    	textList = tree.item(tree.focus())["values"]
    	line = ""
    	for text in textList:
    		if line != "":
    			line += ", " + str(text)
    		else:
    			line += str(text)
    	#master.clipboard_append(line)
    	list_s = line.split()
    	#print(list_s)
    	x=0
    	for l in list_s:
    		book_no=l
    		x +=1
    	k=tkinter.messagebox.askyesno("Delete","Do you want to delete "+line)
    	while k==True:
    		book_no = l
    		result = conn.execute("SELECT * FROM book WHERE book_no='"+book_no.upper()+"'")
    		title=""
    		for x in result:
    			title = x['book_title']
    			print(title)
    		conn.execute("delete from book where book_no='"+book_no.upper()+"'")
    		conn.commit()
    		tkinter.messagebox.showinfo("Success",title+" "+book_no.upper()+" has been deleted")
    		self.display_books()
    		#print(l.upper())
    		break
    	else:
    		pass
    		#print("Sorry")
    def select_student(self,tree,event=None):
    	textList = tree.item(tree.focus())["values"]
    	line = ""
    	for text in textList:
    		if line != "":
    			line += " " + str(text)
    		else:
    			line += str(text)
    	list_s = line.split()
    	for x in range(len(list_s)):
    		x = x

    	#print(list_s[(x-1)])
    	book_no = ""
    	

    	for L in list_s:
    		if(re.search(r'\d', L)):
    			#print("Success")
    			book_no = L
    			break
    	#print(book_no)
    	c.execute("SELECT book_id FROM book WHERE book_no='"+book_no.upper()+"'")
    	file_id = c.fetchone()[0]
    	result = conn.execute("SELECT * FROM borrow WHERE file_id=? AND dob=?",(file_id,list_s[(x-1)]))
    	for res in result:
    		name = res['st_name']
    		id_b = res['id']

    	print(id_b)


    	m = tkinter.messagebox.askyesno("Delete","Do you want to delete record of "+name.upper())
    	while m==True:
    		conn.execute("DELETE FROM borrow WHERE id=?",(id_b,))
    		conn.commit()
    		tkinter.messagebox.showinfo("Success",name+" has been deleted")
    		self.dash()
    		self.all_student()
    		break
    	else:
    		pass



def main():
    try:
        root = Tk()
        app = sms(root)
        app.dash()
        root.bind('<Control-a>', lambda e:app.Add_book_Gui())
        root.bind('<Control-b>', lambda e:app.borrow_book_Gui())
        root.bind('<Control-c>', lambda e:app.clear_st_Gui())

        #app.display_books()
        root.mainloop()
    except:
        pass
if __name__ == '__main__': main()
