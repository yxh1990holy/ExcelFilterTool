import sys
from PySide6.QtWidgets import QApplication
from app.main_window import ExcelFilterWindow


def main():
    app = QApplication(sys.argv)
    window = ExcelFilterWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()