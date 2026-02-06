import sys
from PySide6.QtWidgets import QApplication
from ui import MiniExcelUI

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MiniExcelUI()
    window.show()
    sys.exit(app.exec())