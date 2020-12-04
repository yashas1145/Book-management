'''
Created on Nov 22, 2020

@author: Yashas
'''
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import MySQLdb, sys, datetime
from xlrd import *
from xlsxwriter import *

from PyQt5.uic import loadUiType

ui,_ = loadUiType('Library.ui')
login,_ = loadUiType('Login.ui')

class Login(QWidget, login):
    
    def __init__(self):
        
        QWidget.__init__(self)
        self.setupUi(self)
        self.HandleButton()
        self.DefTheme()
        
        
    def DefTheme(self):
        
        style = open('Themes/darkstyle.css')
        style = style.read()
        self.setStyleSheet(style)
            
    def HandleButton(self):
        
        self.pushButton.clicked.connect(self.HandleLogin)
        self.pushButton_2.clicked.connect(sys.exit)
        
    def HandleLogin(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        uName = self.lineEdit.text()
        uPass = self.lineEdit_2.text()
        
        self.cur.execute('''select uName,uPass from users''')
        data = self.cur.fetchall()
        
        for item in data:
            
            if uName == item[0] and uPass == item[1]:
                
                self.label_4.setText('Welcome!')
                self.LibApp = MainApp()
                self.close()
                self.LibApp.show()
                
            else:
                
                self.label_4.setText('Your credentials might be incorrect...')
                
class MainApp(QMainWindow, ui):
    
    def __init__(self):
        
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_UI()
        self.Handle_Button()
        self.ShowDay2Day()
        self.Show_Category()
        self.Show_Author()
        self.Show_Publisher()
        self.Show_Client()
        self.Show_Books()
        self.Category_Combobox()
        self.Author_Combobox()
        self.Publisher_Combobox() 
        self.darkstyle()
        
    def Handle_UI(self):
        
        self.Close_Themes()
        self.tabWidget.tabBar().setVisible(False)
    
    def Handle_Button(self):
        
        self.pushButton_5.clicked.connect(self.Open_Themes)
        self.pushButton_22.clicked.connect(self.Close_Themes)
        
        self.pushButton.clicked.connect(self.Open_Operations)
        self.pushButton_2.clicked.connect(self.Open_Books)
        self.pushButton_3.clicked.connect(self.Open_Users)
        self.pushButton_4.clicked.connect(self.Open_Settings)
        self.pushButton_26.clicked.connect(self.Open_Clients)
        self.pushButton_36.clicked.connect(self.LogOut)
        
        self.pushButton_8.clicked.connect(self.Day2Day)
        
        self.pushButton_6.clicked.connect(self.Add_Book)
        self.pushButton_10.clicked.connect(self.Search_Book)
        self.pushButton_9.clicked.connect(self.Edit_Book)
        self.pushButton_11.clicked.connect(self.Delete_Book)
        
        self.pushButton_27.clicked.connect(self.Add_Client)
        self.pushButton_28.clicked.connect(self.Search_Client)
        self.pushButton_30.clicked.connect(self.Delete_Client)
        self.pushButton_29.clicked.connect(self.Edit_Client)
        
        self.pushButton_12.clicked.connect(self.Add_User)
        self.pushButton_13.clicked.connect(self.Login_User)
        self.pushButton_25.clicked.connect(self.Edit_User)
        
        self.pushButton_15.clicked.connect(self.Add_Author)
        self.pushButton_16.clicked.connect(self.Add_Category)
        self.pushButton_17.clicked.connect(self.Add_Publisher)
        self.pushButton_31.clicked.connect(self.Delete_Category)
        self.pushButton_33.clicked.connect(self.Delete_Author)
        self.pushButton_32.clicked.connect(self.Delete_Publisher)
        
        self.pushButton_18.clicked.connect(self.darkgray)
        self.pushButton_19.clicked.connect(self.darkstyle)
        self.pushButton_20.clicked.connect(self.light)
        self.pushButton_21.clicked.connect(self.darrkorange)
        
        self.pushButton_7.clicked.connect(self.ExportDay2Day)
        self.pushButton_34.clicked.connect(self.ExportBooks)
        self.pushButton_35.clicked.connect(self.ExportClients)
        
    ##### Button Controls #####
    
    def Open_Themes(self):
        
        self.groupBox_2.show()
    
    def Close_Themes(self):
        
        self.groupBox_2.hide()
    
    def Open_Operations(self):
        
        self.tabWidget.setCurrentIndex(0)
    
    def Open_Books(self):
        
        self.tabWidget.setCurrentIndex(1)
        
    def Open_Clients(self):
        
        self.tabWidget.setCurrentIndex(2)
    
    def Open_Users(self):
        
        self.tabWidget.setCurrentIndex(3)
    
    def Open_Settings(self):
        
        self.tabWidget.setCurrentIndex(4)
        
    def LogOut(self):
        
        self.Window = Login()
        self.Window.show()
        self.close()
    
    ##### Button Controls #####
    
    ##### Day2Day Operations #####
    
    def Day2Day(self):
        
        BookTitle = self.lineEdit.text()
        Type = self.comboBox.currentText()
        DayCount = self.comboBox_2.currentIndex() + 1
        Date = datetime.date.today()
        clName = self.lineEdit_27.text()
        toDate = Date + datetime.timedelta(days=DayCount)
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()

        self.cur.execute('''insert into operations(bName, clName, operationType, DayCount, todayDate, toDate) values (%s,%s,%s,%s,%s,%s)''',(BookTitle, clName, Type, DayCount, Date, toDate))
        self.con.commit()
        self.statusBar().showMessage('Operation noted')
        self.ShowDay2Day()
       
    def ShowDay2Day(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select bName, clName, operationType, todayDate, toDate from operations''')
        data = self.cur.fetchall()
        
        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)
        for row, form in enumerate(data):
            for col, item in enumerate(form):
                self.tableWidget.setItem(row, col, QTableWidgetItem(str(item)))
                col += 1
            rowPos = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPos)
                
    
    ##### Day2Day Operations #####
    
    ##### Books Operations #####
    
    def Show_Books(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select bName, bDesc, bCode, bPrice from books''')
        data_books = self.cur.fetchall()
        
        if data_books:
            self.tableWidget_5.setRowCount(0)
            self.tableWidget_5.insertRow(0)
            for row, form in enumerate(data_books):
                for col, item in enumerate(form):
                    self.tableWidget_5.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_pos = self.tableWidget_5.rowCount()
                self.tableWidget_5.insertRow(row_pos)
    
    def Add_Book(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor() 
        
        BookTitle = self.lineEdit_2.text()
        BookDesc = self.textEdit.toPlainText()
        BookID = self.lineEdit_3.text()
        BookAuthor = self.comboBox_3.currentIndex()
        BookPublisher = self.comboBox_4.currentIndex()
        BookCategory = self.comboBox_5.currentIndex()
        BookPrice = self.lineEdit_7.text()

        self.cur.execute('''insert into books (bName, bDesc, bCode, bCategory, bAuthor, bPublisher, bPrice) values (%s, %s, %s, %s, %s, %s, %s)''', (BookTitle, BookDesc, BookID, BookCategory, BookAuthor, BookPublisher, BookPrice))
        self.con.commit()
      
        self.lineEdit_2.setText('')
        self.lineEdit_3.setText('')
        self.textEdit.setText('')
        self.lineEdit_7.setText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)
        self.Show_Books()
        
        self.statusBar().showMessage('New book added')
        
    def Delete_Book(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        bookNameForSearch = self.lineEdit_6.text()
        
        warn = QMessageBox.warning(self, 'Confirm Deletion', 'Are you sure you want to delete?', QMessageBox.Yes | QMessageBox.No)
        
        if warn == QMessageBox.Yes:
            
            self.cur.execute('''delete from books where bName=%s''',[bookNameForSearch])
            self.con.commit()
            self.statusBar().showMessage('Book deleted')
            self.lineEdit_6.setText('')
            self.lineEdit_4.setText('')
            self.lineEdit_8.setText('')
            self.textEdit_2.setPlainText('')
            self.comboBox_8.setCurrentIndex(0)
            self.comboBox_7.setCurrentIndex(0)
            self.comboBox_6.setCurrentIndex(0)
            self.Show_Books()
        
    
    def Search_Book(self):
        
        self.statusBar().showMessage('')
        
        bookNameForSearch = self.lineEdit_5.text()
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        try:
            
            self.cur.execute('''select * from books where bName = %s''',[bookNameForSearch])
            self.con.commit()
            
            data = self.cur.fetchone()
        
            self.lineEdit_6.setText(data[1])
            self.lineEdit_4.setText(data[3])
            self.lineEdit_8.setText(str(data[7]))
            self.textEdit_2.setPlainText(data[2])
            self.comboBox_8.setCurrentIndex(data[4])
            self.comboBox_7.setCurrentIndex(data[5])
            self.comboBox_6.setCurrentIndex(data[6])
            
            self.lineEdit_5.setText('')
            
        except:
            
            self.statusBar().showMessage('Book not found')
        
    
    def Edit_Book(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        bookNameForSearch = self.lineEdit_6.text()
        BookTitle = self.lineEdit_6.text()
        BookDesc = self.textEdit_2.toPlainText()
        BookID = self.lineEdit_4.text()
        BookAuthor = self.comboBox_8.currentIndex()
        BookPublisher = self.comboBox_7.currentIndex()
        BookCategory = self.comboBox_6.currentIndex()
        BookPrice = self.lineEdit_8.text()
        
        self.cur.execute('''update books set bName=%s, bDesc=%s, bCode=%s, bAuthor=%s, bPublisher=%s, bCategory=%s, bPrice=%s where bName=%s''',(BookTitle,BookDesc,BookID,BookAuthor,BookPublisher,BookCategory,BookPrice,bookNameForSearch))
        self.con.commit()
        
        self.statusBar().showMessage('Book details updated')
        self.Show_Books()
        
        self.lineEdit_6.setText('')
        self.lineEdit_4.setText('')
        self.lineEdit_8.setText('')
        self.textEdit_2.setPlainText('')
        self.comboBox_8.setCurrentIndex(0)
        self.comboBox_7.setCurrentIndex(0)
        self.comboBox_6.setCurrentIndex(0)
    
    def Author_Combobox(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select aName from authors''')
        data_author = self.cur.fetchall()
        
        self.comboBox_3.clear()
        
        self.comboBox_3.addItem('-----Select Author-----')
        self.comboBox_8.addItem('-----Author-----')
        
        for i in data_author:
            self.comboBox_3.addItem(i[0])
            self.comboBox_8.addItem(i[0])
            
    def Category_Combobox(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select cName from categories''')
        data_category = self.cur.fetchall()
        
        self.comboBox_5.clear()
        
        self.comboBox_5.addItem('-----Select Category-----')
        self.comboBox_6.addItem('-----Category-----')
        
        for i in data_category:
            self.comboBox_5.addItem(i[0])
            self.comboBox_6.addItem(i[0])
            
    
    def Publisher_Combobox(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select pName from publishers''')
        data_publisher = self.cur.fetchall()
        
        self.comboBox_4.clear()
        
        self.comboBox_4.addItem('-----Select Publisher-----')
        self.comboBox_7.addItem('-----Publisher-----')
        
        for i in data_publisher:
            self.comboBox_4.addItem(i[0])
            self.comboBox_7.addItem(i[0])
    
    ##### Books Operations #####
    
    ##### Users Opereations #####
    
    def Add_User(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        uName = self.lineEdit_11.text()
        uPass = self.lineEdit_10.text()
        uMail = self.lineEdit_9.text()
        
        self.cur.execute('''insert into users (uName, uMail, uPass) values (%s, %s, %s)''',(uName, uMail, uPass))
        self.con.commit()
        
        self.statusBar().showMessage('New user added')
    
    def Login_User(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        uName = self.lineEdit_13.text()
        uPass = self.lineEdit_12.text()
        
        self.cur.execute('''select uName,uMail,uPass from users''')
        data = self.cur.fetchall()
        
        for item in data:
            
            if uName == item[0] and uPass == item[2]:
                
                self.statusBar().showMessage('User log in success')
                self.groupBox_4.setEnabled(True)
                
                self.lineEdit_15.setText(uName)
                self.lineEdit_14.setText(uPass)
                self.lineEdit_16.setText(item[1])
                
                break
            
            else:
                
                self.statusBar().showMessage('Incorrect credentials')
                
    def Edit_User(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        InitName = self.lineEdit_13.text()
        uName = self.lineEdit_15.text()
        uPass = self.lineEdit_14.text() 
        uMail = self.lineEdit_16.text()
        
        self.cur.execute('''update users set uName=%s, uMail=%s, uPass=%s where uName=%s''',(uName,uMail,uPass,InitName))
        self.con.commit()
        self.statusBar().showMessage('User details updated')
    
    ##### Users Opereations #####
    
    ##### Client Opereations #####
    
    def Show_Client(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select clName, clMail, clNational from clients''')
        data_client = self.cur.fetchall()
        
        if data_client:
            self.tableWidget_6.setRowCount(0)
            self.tableWidget_6.insertRow(0)
            for row, form in enumerate(data_client):
                for col, item in enumerate(form):
                    self.tableWidget_6.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_pos = self.tableWidget_6.rowCount()
                self.tableWidget_6.insertRow(row_pos)
    
    def Add_Client(self):
        
        clName = self.lineEdit_22.text()
        clMail = self.lineEdit_20.text()
        clNID = self.lineEdit_21.text()
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''insert into clients (clName, clMail, clNational) values (%s,%s,%s)''',(clName,clMail,clNID))
        self.con.commit()
        self.Show_Client()
        
        self.statusBar().showMessage('New client added')
        
    def Search_Client(self):
        
        clName = self.lineEdit_23.text()
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        try:
            
            self.cur.execute('''select clName, clMail, clNational from clients where clName=%s''',[clName])
            data = self.cur.fetchone()
        
            self.lineEdit_24.setText(data[0])
            self.lineEdit_26.setText(data[1])
            self.lineEdit_25.setText(data[2])
            
        except:
            
            self.statusBar().showMessage('Client not found')
    
    def Edit_Client(self):
        
        clName = self.lineEdit_23.text()
        
        name = self.lineEdit_24.text()
        mail = self.lineEdit_26.text()
        national = self.lineEdit_25.text()
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''update clients set clName=%s, clMail=%s, clNational=%s where clName=%s''',(name,mail,national,clName))
        self.con.commit()
        
        self.statusBar().showMessage('Client details updated')
        self.Show_Client()
    
    def Delete_Client(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        clName = self.lineEdit_23.text()
        
        warn = QMessageBox.warning(self, 'Confirm Deletion', 'Are you sure you want to delete?', QMessageBox.Yes | QMessageBox.No)
        
        if warn == QMessageBox.Yes:
            
            self.cur.execute('''delete from clients where clName=%s''',[clName])
            self.con.commit()
            self.Show_Client()
            self.statusBar().showMessage('Client deleted')
            self.lineEdit_23.setText('')
            self.lineEdit_24.setText('')
            self.lineEdit_25.setText('')
            self.lineEdit_26.setText('')
    
    ##### Client Opereations #####
    
    ##### Settings Operations #####
    
    def Add_Category(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        CategoryName = self.lineEdit_18.text()
        
        self.cur.execute('''insert into categories (cName) values (%s)''',(CategoryName,))
        self.con.commit()
        
        self.statusBar().showMessage('New category added')
        self.lineEdit_18.setText('')
        self.Show_Category()
        self.Category_Combobox()
        
    def Delete_Category(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        CategoryName = self.lineEdit_18.text()
        
        self.cur.execute('''delete from categories where cName=%s''',[CategoryName])
        self.con.commit()
        self.statusBar().showMessage('Category deleted')
        self.lineEdit_18.setText('')
        self.Show_Category()
        self.Category_Combobox()
        
    def Show_Category(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select cName from categories''')
        data_category = self.cur.fetchall()
        
        if data_category:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row, form in enumerate(data_category):
                for col, item in enumerate(form):
                    self.tableWidget_3.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_pos = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_pos)
            
    def Add_Author(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        AuthorName = self.lineEdit_17.text()
        
        self.cur.execute('''insert into authors (aName) values (%s)''',(AuthorName,))
        self.con.commit()
        
        self.statusBar().showMessage('New author added')
        self.lineEdit_17.setText('')
        self.Show_Author()
        self.Author_Combobox()
        
    def Delete_Author(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        AuthorName = self.lineEdit_17.text()
        
        self.cur.execute('''delete from authors where aName=%s''',[AuthorName])
        self.con.commit()
        self.statusBar().showMessage('Author deleted')
        self.lineEdit_17.setText('')
        self.Show_Author()
        self.Author_Combobox()
        
    def Show_Author(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select aName from authors''')
        data_author = self.cur.fetchall()
        
        if data_author:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, form in enumerate(data_author):
                for col, item in enumerate(form):
                    self.tableWidget_2.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_pos = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_pos)
    
    def Add_Publisher(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        PublisbherName = self.lineEdit_19.text()
        
        self.cur.execute('''insert into publishers (pName) values (%s)''',(PublisbherName,))
        self.con.commit()
        
        self.statusBar().showMessage('New publisher added')
        self.lineEdit_19.setText('')
        self.Show_Publisher()
        self.Publisher_Combobox()
        
    def Delete_Publisher(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        PublisherName = self.lineEdit_19.text()
        
        self.cur.execute('''delete from publishers where pName=%s''',[PublisherName])
        self.con.commit()
        self.statusBar().showMessage('Publisher deleted')
        self.lineEdit_19.setText('')
        self.Show_Publisher()
        self.Publisher_Combobox()
        
    def Show_Publisher(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select pName from publishers''')
        data_publisher = self.cur.fetchall()
        
        if data_publisher:
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)
            for row, form in enumerate(data_publisher):
                for col, item in enumerate(form):
                    self.tableWidget_4.setItem(row, col, QTableWidgetItem(str(item)))
                    col += 1
                row_pos = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_pos)
    
    ##### Settings Operations #####
    
    ##### Theme Operations #####
    
    def darkgray(self):
        
        style = open('Themes/darkgray.css')
        style = style.read()
        self.setStyleSheet(style)
    
    def darkstyle(self):
        
        style = open('Themes/darkstyle.css')
        style = style.read()
        self.setStyleSheet(style)
    
    def light(self):
        
        style = open('Themes/light.css')
        style = style.read()
        self.setStyleSheet(style)
        
    def darrkorange(self):
        
        style = open('Themes/darkorange.css')
        style = style.read()
        self.setStyleSheet(style)
    
    ##### Theme Operations #####
    
    ##### Export Operations #####
    
    def ExportDay2Day(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select bName, clName, operationType, todayDate, toDate from operations''')
        data = self.cur.fetchall()
        
        NewExcel = Workbook('Operations.xlsx')
        Sheet1 = NewExcel.add_worksheet()
        
        Sheet1.write(0,0,'Book Name')
        Sheet1.write(0,1,'Client Name')
        Sheet1.write(0,2,'Return / Borrow')
        Sheet1.write(0,3,'From')
        Sheet1.write(0,4,'To')
        
        rowPos = 1
        for row in data:
            colPos = 0
            for item in row:
                Sheet1.write(rowPos, colPos, str(item))
                colPos += 1 
            rowPos += 1
        
        self.statusBar().showMessage('Day operations exported')
        NewExcel.close()
    
    def ExportBooks(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select bName, bDesc, bCode, bPrice from books''')
        data = self.cur.fetchall()
        
        NewExcel = Workbook('Books.xlsx')
        Sheet1 = NewExcel.add_worksheet()
        
        Sheet1.write(0,0,'Book name')
        Sheet1.write(0,1,'Book desc')
        Sheet1.write(0,2,'Book code')
        Sheet1.write(0,3,'Book price')
        
        rowPos = 1
        for row in data:
            colPos = 0
            for item in row:
                Sheet1.write(rowPos, colPos, str(item))
                colPos += 1 
            rowPos += 1
        
        self.statusBar().showMessage('Book details exported')
        NewExcel.close()
    
    def ExportClients(self):
        
        self.con = MySQLdb.connect('localhost', 'root', 'drowssap', 'Library')
        self.cur = self.con.cursor()
        
        self.cur.execute('''select clName, clMail, clNational from clients''')
        data = self.cur.fetchall()
        
        NewExcel = Workbook('Clients.xlsx')
        Sheet1 = NewExcel.add_worksheet()
        
        Sheet1.write(0,0,'Client name')
        Sheet1.write(0,1,'Client mail')
        Sheet1.write(0,2,'Client national ID')
        
        rowPos = 1
        for row in data:
            colPos = 0
            for item in row:
                Sheet1.write(rowPos, colPos, str(item))
                colPos += 1 
            rowPos += 1
        
        self.statusBar().showMessage('Client details exported')
        NewExcel.close()
    
    ##### Export Operations #####
        
def main():
    
    app = QApplication(sys.argv)
    
    Window = Login()
    Window.show()
    app.exec_()
    
if __name__ == '__main__':
    
    main()