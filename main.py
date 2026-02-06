import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAbstractItemView)
from formula_engine import FormulaEngine
from PySide6.QtCore import Qt

ROWS = 60
COLS = 30

class MiniExcel(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mini Excel - Qt Edition")
        self.resize(1200, 700)

        self.table = QTableWidget(ROWS, COLS)
        self.setCentralWidget(self.table)

        self.setup_header()
        self.setup_table_behavior()

        self.formula_engine = FormulaEngine(self.table)
        self.table.itemChanged.connect(self.formula_engine.process_item)
        self.table.itemDoubleClicked.connect(self.on_edit_start)
        self.table.itemDelegate().closeEditor.connect(self.on_edit_end)

    def setup_header(self):
        # Sütun harfleri
        for c in range(COLS):
            self.table.setHorizontalHeaderItem(c, QTableWidgetItem(chr(ord("A") + c)))
            # Satır numaraları
            for r in range(ROWS):
                self.table.setVerticalHeaderItem(r, QTableWidgetItem(str(r+1)))
    
    def setup_table_behavior(self):
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)

        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(24)

        self.table.setShowGrid(True)
        self.table.setAlternatingRowColors(False)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
    
    def on_edit_start(self, item):
        formula = item.data(Qt.UserRole)
        if formula:
            self.table.blockSignals(True)
            item.setText(formula)
            self.table.blockSignals(False)

    def on_edit_end(self, editor=None, hint=None):
        item = self.table.currentItem()
        if not item:
            return

        formula = item.data(Qt.UserRole)
        if not formula:
            return

        self.table.blockSignals(True)
        value = self.formula_engine.evaluate(formula)
        item.setText(value)
        self.table.blockSignals(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MiniExcel()
    window.show()
    sys.exit(app.exec())