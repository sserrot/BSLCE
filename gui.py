# -*- coding: utf-8 -*-
"""
GUI.py
Create GUI for checking attendance and sending out emails
$ init_ui
$ attendance_on_click
$ master_on_click
$ run_program
"""

import sys
import threading
from PyQt5.QtWidgets import QLineEdit, QPushButton, QFileDialog, QApplication, QMainWindow, QTextEdit, QLabel, QPlainTextEdit
import find_non_attendants
import excel_parse
import send_emails
import find_path # find a better way to do this



class App(QMainWindow):
    sender_address = ''
    password = ''
    subject = ''
    body_text = ''
    attended_ID = []
    master_information = []
    firstURL = ''

    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        self.statusBar()

        # User Email Address
        self.address_label = QLabel(self)
        self.address_label.setText("Email Address")
        self.address_label.move(145, 120)

        self.address_text = QLineEdit(self)
        self.address_text.move(250, 120)

        # User password
        self.pass_label = QLabel(self)
        self.pass_label.setText("Password")
        self.pass_label.move(145, 160)

        self.pass_text = QLineEdit(self)
        self.pass_text.setEchoMode(QLineEdit.Password)
        self.pass_text.move(250, 160)

        # Subject of Email
        self.subject_label = QLabel(self)
        self.subject_label.setText("Email Subject")
        self.subject_label.move(145, 200)

        self.subject_text = QLineEdit(self)
        self.subject_text.move(250, 200)

        # Body of Email
        self.body_label = QLabel(self)
        self.body_label.setText("Email Body")
        self.body_label.move(145, 240)

        self.body_text = QPlainTextEdit(self)
        self.body_text.setMinimumHeight(150)
        self.body_text.setMinimumWidth(250)
        self.body_text.move(250, 240)

        # Attendance excel file
        att_button = QPushButton('Attendance File', self)
        att_button.setToolTip('Open Attendance File')
        att_button.move(145, 400)
        att_button.clicked.connect(self.attendance_on_click)

        self.att_text = QTextEdit(self)
        self.att_text.setReadOnly(True)
        self.att_text.move(250, 400)

        # Master excel file
        mast_button = QPushButton('Master File', self)
        mast_button.setToolTip('Open Master File')
        mast_button.move(145, 460)
        mast_button.clicked.connect(self.master_on_click)

        self.mast_text = QTextEdit(self)
        self.mast_text.setReadOnly(True)
        self.mast_text.move(250, 460)

        # run Program
        non_att_button = QPushButton('Run Program', self)
        non_att_button.setToolTip('Open Master File')
        non_att_button.move(250, 500)
        non_att_button.clicked.connect(self.run_program)

        self.setGeometry(300, 300, 600, 600)
        self.setWindowTitle('Check Attendance')
        self.show()


    def attendance_on_click(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file')
        filename = find_path.pathsplit(fname[0])
        App.attended_ID = excel_parse.read_attendance(fname[0])
        self.att_text.setText(filename)


    def master_on_click(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file')
        filename = find_path.pathsplit(fname[0])
        App.master_information = excel_parse.read_master(fname[0])
        self.mast_text.setText(filename)

    def populate_fields(self):
        App.sender_address = self.address_text.text()
        App.password = self.pass_text.text()
        App.subject = self. subject_text.text()
        App.body_text = self.body_text.toPlainText()

    def run_program(self):
        self.populate_fields()
        non_attendants = find_non_attendants.main(App.master_information, App.attended_ID)
        email_thread = threading.Thread(target=send_emails.email(sender_address=App.sender_address, password=App.password, subject=App.subject, body_text=App.body_text, non_attendant_students=non_attendants))
        email_thread.start()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())