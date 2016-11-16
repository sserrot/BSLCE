"""

"""

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import pyqtSlot
import read_attendance_file
import read_master_file
import find_non_attendants
import send_emails


class App(QMainWindow):
    sender_address = ''
    password = ''
    subject = ''
    text_body = ''
    attended_ID = []
    master_information = []

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.statusBar()

        # User Email Address
        self.address_text = QLineEdit(self)
        self.address_text.move(30, 110)

        # User password
        self.pass_text = QLineEdit(self)
        self.pass_text.setEchoMode(QLineEdit.Password)
        self.pass_text.move(60, 110)
        # Subject of Email
        self.subject = QLineEdit(self)
        self.subject.move(90, 110)
        # Body of Email
        self.subject = QLineEdit(self)
        self.subject.move(100, 110)

        # Attendance excel file
        att_button = QPushButton('Attendance File', self)
        att_button.setToolTip('Open Attendance File')
        att_button.move(130, 180)
        att_button.clicked.connect(self.attendance_on_click)

        self.att_text = QLineEdit(self)
        self.att_text.move(130, 210)

        # Master excel file
        mast_button = QPushButton('Master File', self)
        mast_button.setToolTip('Open Master File')
        mast_button.move(130, 240)
        mast_button.clicked.connect(self.master_on_click)

        self.mast_text = QLineEdit(self)
        self.mast_text.move(130, 270)

        # run Program
        non_att_button = QPushButton('Run Program', self)
        non_att_button.setToolTip('Open Master File')
        non_att_button.move(130, 330)
        non_att_button.clicked.connect(self.non_attendants_on_click)

        self.setGeometry(300, 300, 600, 600)
        self.setWindowTitle('Check Attendance')
        self.show()

    @pyqtSlot()
    def attendance_on_click(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')
        App.attended_ID = read_attendance_file.attended(fname[0])
        self.att_text.setText(str(fname[0]))

    @pyqtSlot()
    def master_on_click(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')
        App.master_information = read_master_file.master(fname[0])
        self.mast_text.setText(str(fname[0]))

    @pyqtSlot()
    def non_attendants_on_click(self):
        App.sender_address = self.address_text.text()
        non_attendants = find_non_attendants.main(App.master_information, App.attended_ID)
        send_emails.email(sender_address=App.sender_address, password=App.password, subject=App.subject,
                          text_body=App.text_body, non_attendant_students=non_attendants)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())