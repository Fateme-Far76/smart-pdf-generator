from PyQt5.QtWidgets import QApplication
import sys
from ui_app import ExcelFilterApp

def main():
    app = QApplication(sys.argv)
    window = ExcelFilterApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
