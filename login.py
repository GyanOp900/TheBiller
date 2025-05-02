from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox
from PyQt6.QtGui import QIcon, QFont
import sys
import hashlib
import datetime
from PyQt6.QtCore import (
    Qt, QTimer
)

class LoginWindow(QWidget):

    


    hashsafepass = 'd8386cefb5c76076bf0ce4b63ec8b75667cbe244e6db5b69114030d2b0bdb920'


    def __init__(self, app):
        super().__init__()
        self.setWindowTitle("Paid Login (TheBiller V1 ~THE QR UPDATE~)")
        self.app = app
        self.setGeometry(100,100,400,250)
        self.setWindowIcon(QIcon("logo.ico"))

        layout = QVBoxLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Username")
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_input)

        login_button = QPushButton("Login")
        login_button.clicked.connect(self.check_login)
        layout.addWidget(login_button)

        self.time_label = QLabel(datetime.datetime.now().strftime("%H:%M:%S"))
        font = QFont("Garamond", 20)
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_label.setFont(font)
        layout.addWidget(self.time_label)

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        self.update_time()

        self.setStyleSheet("""
                            QPushButton {
                        padding: 10px;
                           }
                           
                           QLineEdit {
                        padding: 10px   
                           }
                            """)

        self.setLayout(layout)

    def update_time(self):
        self.time_label.setText(datetime.datetime.now().strftime("%H:%M:%S"))
        
    def check_login(self):
        username = self.username_input.text()
        password = self.password_input.text()


         
        

        if LoginWindow.hashsafepass == hashlib.sha256(password.encode()).hexdigest():
            self.open_billing_app()
            self.close()


        else:
            QMessageBox.warning(self, "Error", "Invalid Username or Password")

    def open_billing_app(self):
        from main import BillingSystem
        self.billing_app = BillingSystem(self.app)
        self.billing_app.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    login = LoginWindow(app=app)
    login.show()
    sys.exit(app.exec())


