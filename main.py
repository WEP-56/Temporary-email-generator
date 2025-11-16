import sys
from email_gui import QApplication, EmailRegisterApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EmailRegisterApp()
    window.show()
    sys.exit(app.exec())