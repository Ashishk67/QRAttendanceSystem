from __future__ import print_function
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QTextEdit
import sys
import webbrowser
from oauth2client.service_account import ServiceAccountCredentials
from pyzbar.pyzbar import decode
from PIL import Image
import os.path
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from PyQt5.QtWidgets import QPlainTextEdit
from PyQt5 import QtGui
from PyQt5.QtWidgets import QApplication,QWidget, QVBoxLayout, QPushButton, QFileDialog , QLabel, QTextEdit


class Window(QWidget):
    def __init__(self):
        super().__init__()

        self.title = "QR Attendance"
        self.top = 300
        self.left = 300
        self.width = 600
        self.height = 600
        self.InitWindow()


    def InitWindow(self):
        self.setWindowIcon(QtGui.QIcon("C:/Users/Ashish/qricon.ico"))
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        vbox = QVBoxLayout()
        self.btn1 = QPushButton("Mark Attendance")
        self.btn1.clicked.connect(self.scan)
        vbox.addWidget(self.btn1)
        self.btn2 = QPushButton("Check Attendance History")
        self.btn2.clicked.connect(self.getImage)
        vbox.addWidget(self.btn2)
        self.btn3 = QPushButton("Exit")
        self.btn3.clicked.connect(self.close)
        vbox.addWidget(self.btn3)
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        vbox.addWidget(self.text_area)
        self.setLayout(vbox)
        self.show()

    def close(self):
        sys.exit(0)

    def scan(self):
        try:
            self.text_area.clear()
            fname = QFileDialog.getOpenFileName(self,'Select QR image','C:/Users/Ashish/Desktop/qr codes', "Image files (*.jpg *.png)")
            imagePath = fname[0]
            x = str(decode(Image.open(imagePath)))
            url=x.split("'")[1]
            webbrowser.open(url, new=0, autoraise=True)
        except:
            self.text_area.clear()
            self.text_area.append("Please select a valid QR image file!")
    def getImage(self):
        try:
            self.text_area.clear()
            fname = QFileDialog.getOpenFileName(self,'Select QR image','C:/Users/Ashish/Desktop/qr codes', "Image files (*.jpg *.png)")
            imagePath = fname[0]
            x = str(decode(Image.open(imagePath)))
            url=x.split("'")[1]
            name=url.split("=")[2]
            SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
            SPREADSHEET_ID = '14eFVHr0tcm6540gzIJ5SLdSOaFeXbPfs3tQNeg7erBI'
            RANGE_NAME = 'Responses!A2:B1000'

            creds = None
            creds = service_account.Credentials.from_service_account_file('credentials.json',scopes=SCOPES)
            service = build('sheets', 'v4', credentials=creds)
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,range=RANGE_NAME).execute()
            values=result.get('values',[])
            values_list=[]
            s=f'Attendance record for {name}\nDay(s) present:'
            self.text_area.append(s)
            for i in range(len(values)):
                if values[i][1]==name:
                    values_list.append(values[i])
            for i in values_list:
                self.text_area.append(i[0])
            self.text_area.append(f'\nTotal day(s) present : {len(values_list)}\n')
        except:
            self.text_area.clear()
            self.text_area.append("Please select a valid QR image.")
App = QApplication(sys.argv)
window = Window()
sys.exit(App.exec())