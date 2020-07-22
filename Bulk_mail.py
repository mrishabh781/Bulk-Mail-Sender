
import smtplib
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from PyQt5.QtGui import QIcon
import pandas as pd



class Ui_MainWindow(object):
    login_name = ''
    login_password = ''
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1059, 883)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 70, 111, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(0, 110, 67, 20))
        self.label_2.setObjectName("label_2")
        self.loginid = QtWidgets.QLineEdit(self.centralwidget)
        self.loginid.setGeometry(QtCore.QRect(130, 70, 291, 25))
        self.loginid.setInputMethodHints(QtCore.Qt.ImhEmailCharactersOnly)
        self.loginid.setObjectName("loginid")
        #Ui_MainWindow.login_name = self.loginid.text()
        self.password = QtWidgets.QLineEdit(self.centralwidget)
        self.password.setGeometry(QtCore.QRect(130, 110, 291, 25))
        self.password.setAutoFillBackground(False)
        self.password.setObjectName("password")
        self.password.setEchoMode(QLineEdit.Password)
        #Ui_MainWindow.login_password = self.password.text()
        self.login = QtWidgets.QPushButton(self.centralwidget)
        self.login.setGeometry(QtCore.QRect(130, 150, 121, 41))
        self.login.setObjectName("login")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(440, 210, 89, 25))
        self.pushButton_2.setObjectName("pushButton_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(0, 200, 101, 41))
        self.label_3.setObjectName("label_3")
        self.path_to_excel = QtWidgets.QLineEdit(self.centralwidget)
        self.path_to_excel.setGeometry(QtCore.QRect(130, 210, 291, 25))
        self.path_to_excel.setObjectName("path_to_excel")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(0, 250, 91, 31))
        self.label_4.setObjectName("label_4")
        self.sender_column = QtWidgets.QComboBox(self.centralwidget)
        self.sender_column.setGeometry(QtCore.QRect(130, 250, 291, 25))
        self.sender_column.setObjectName("sender_column")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(130, 0, 251, 61))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(0, 310, 67, 17))
        self.label_6.setObjectName("label_6")
        self.subject = QtWidgets.QLineEdit(self.centralwidget)
        self.subject.setGeometry(QtCore.QRect(130, 300, 711, 25))
        self.subject.setObjectName("subject")

        self.mail_body = QtWidgets.QTextEdit(self.centralwidget)
        self.mail_body.setGeometry(QtCore.QRect(130, 350, 721, 251))
        self.mail_body.setObjectName("mail_body")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(0, 600, 151, 51))
        self.label_8.setObjectName("label_8")
        self.column_selector = QtWidgets.QComboBox(self.centralwidget)
        self.column_selector.setGeometry(QtCore.QRect(130, 610, 511, 31))
        self.column_selector.setObjectName("column_selector")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(0, 360, 67, 17))
        self.label_9.setObjectName("label_9")
        self.add_column = QtWidgets.QPushButton(self.centralwidget)
        self.add_column.setGeometry(QtCore.QRect(670, 610, 121, 31))
        self.add_column.setObjectName("add_column")

        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(0, 710, 131, 41))
        self.label_10.setObjectName("label_10")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(530, 780, 411, 31))
        self.progressBar.setProperty("value", 0)    
        self.progressBar.setObjectName("progressBar")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(960, 770, 89, 41))
        self.pushButton_3.setObjectName("pushButton_3")
        self.preview = QtWidgets.QPushButton(self.centralwidget)
        self.preview.setGeometry(QtCore.QRect(540, 690, 111, 41))
        self.preview.setObjectName("preview")
        self.send = QtWidgets.QPushButton(self.centralwidget)
        self.send.setGeometry(QtCore.QRect(770, 690, 131, 41))
        self.send.setObjectName("send")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(570, 210, 89, 25))
        self.pushButton.setObjectName("pushButton")
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setGeometry(QtCore.QRect(140, 670, 381, 151))
        self.plainTextEdit.setReadOnly(True)
        self.plainTextEdit.setObjectName("plainTextEdit")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1059, 22))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuAbout = QtWidgets.QMenu(self.menubar)
        self.menuAbout.setObjectName("menuAbout")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.toolBar = QtWidgets.QToolBar(MainWindow)
        self.toolBar.setObjectName("toolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.toolBar)
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.actionUsing_the_software = QtWidgets.QAction(MainWindow)
        self.actionUsing_the_software.setObjectName("actionUsing_the_software")
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionUsing_the_software)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menubar.addAction(self.menuAbout.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Login E-mail ID"))
        self.label_2.setText(_translate("MainWindow", "Password"))
        self.login.setText(_translate("MainWindow", "Login"))
        self.login.clicked.connect(self.pressed)
        self.pushButton_2.setText(_translate("MainWindow", "Browse"))
        self.pushButton_2.clicked.connect(self.file_selector)
        self.label_3.setText(_translate("MainWindow", "Path To Excel"))
        self.label_4.setText(_translate("MainWindow", "Sender Email"))
        self.label_5.setText(_translate("MainWindow", "Please Login To Your Gmail Account"))
        self.label_6.setText(_translate("MainWindow", "Subject"))
        self.label_8.setText(_translate("MainWindow", "Column Name"))
        self.label_9.setText(_translate("MainWindow", "Body"))
        self.add_column.setText(_translate("MainWindow", "Add"))
        self.add_column.clicked.connect(self.add)
        self.label_10.setText(_translate("MainWindow", "Selected Columns"))
        self.pushButton_3.setText(_translate("MainWindow", "Exit"))
        self.preview.setText(_translate("MainWindow", "Preview"))
        self.send.setText(_translate("MainWindow", "Send"))
        self.send.clicked.connect(self.send_mail)
        self.pushButton.setText(_translate("MainWindow", "Load"))
        self.pushButton.clicked.connect(self.load_file)
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.menuAbout.setTitle(_translate("MainWindow", "About"))
        self.toolBar.setWindowTitle(_translate("MainWindow", "toolBar"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.actionUsing_the_software.setText(_translate("MainWindow", "Using the software"))

    def pressed(self):
        Ui_MainWindow.login_name = self.loginid.text()
        Ui_MainWindow.login_password = self.password.text()

        x=self.login_mail(Ui_MainWindow.login_name,Ui_MainWindow.login_password)
        self.label_5.setText(x[0])
        #self.label_5.adjustSize()
        self.label_5.setStyleSheet("background-color:{};".format(x[1]))

    def file_selector(self):
        self.file_name = QFileDialog.getOpenFileName()
        print(self.file_name)
        self.path_to_excel.setText(self.file_name[0])

    def load_file(self):
        self.plainTextEdit.setPlainText('')
        self.content = []
        self.num = 0
        file = self.path_to_excel.text()
        self.data = pd.read_excel(file)
        self.heading = []
        for i in self.data:
            self.heading.append(i)
        self.sender_column.clear()
        self.sender_column.addItems(self.heading)
        self.column_selector.clear()
        self.column_selector.addItems(self.heading)


    def add(self):
        self.content.append(self.column_selector.currentText())
        self.num +=1
        text = self.plainTextEdit.toPlainText()
        text += str(self.num)+'. '+self.column_selector.currentText()+'\n'
        self.plainTextEdit.setPlainText(text)


    def send_mail(self):
        self.body_text = self.mail_body.toPlainText()
        print(self.body_text)
        self.message_body = self.body_text.split('{}')
        body_txt= self.message_constructor()
        count = 0
        to_id = list(self.data[self.sender_column.currentText()])
        total = len(self.data[self.sender_column.currentText()])
        for i,j in zip(to_id,body_txt):
            self.send_email(Ui_MainWindow.login_name,i,j)
            count +=1
            self.progressBar.setProperty("value", (count/total)*100)

    def message_constructor(self):
        body_txt = []
        for k in range(len(self.content)):
            for i in range(len(list(self.data[self.content[0]]))):
                if k == 0:
                    body_txt.append(self.message_body[k] + list(self.data[self.content[k]])[i])
                else:
                    body_txt[i] += self.message_body[k] + list(self.data[self.content[k]])[i]
                if k == len(self.content)-1:
                    body_txt[i] += self.message_body[k+1]

        return(body_txt)


    def login_mail(self,user_name,password):
        self.flag = 0    
        try:
            self.email = smtplib.SMTP('smtp.gmail.com', 587)
            self.email.starttls()
            self.email.login(user_name,password)
            self.flag = 1
            return ("Login Successful!",'green')
        except:
            return ("Login failed. Please try again.",'red')

    def send_email(self,user_name,email_id,message):
        if self.flag ==1:
            x = self.email.sendmail(user_name, email_id, message)
            print(x)
        else:
            print("Please login first.")





if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

