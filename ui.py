from PySide6.QtWidgets import (
    QMainWindow,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView,
    QWidget,
    QVBoxLayout,
    QLineEdit
)
from PySide6.QtCore import Qt

from formula_engine import FormulaEngine

ROWS = 60
COLS = 30


class MiniExcelUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Mini Excel – AST Engine")
        self.resize(1200, 700)

        # ===============================
        # CENTRAL WIDGET + LAYOUT
        # ===============================
        central = QWidget(self)
        self.setCentralWidget(central)

        layout = QVBoxLayout()
        central.setLayout(layout)

        # ===============================
        # FORMULA BAR
        # ===============================
        self.formula_bar = QLineEdit()
        self.formula_bar.setPlaceholderText("Formula")
        self.formula_bar.setFixedHeight(28)
        layout.addWidget(self.formula_bar)

        # ===============================
        # TABLE
        # ===============================
        self.table = QTableWidget(ROWS, COLS)
        layout.addWidget(self.table)

        self._setup_headers()
        self._setup_table_behavior()

        # ===============================
        # ENGINE
        # ===============================
        self.engine = FormulaEngine(self.table)

        # ===============================
        # SIGNALS
        # ===============================
        self.table.itemChanged.connect(self.engine.process_item)
        self.table.currentItemChanged.connect(self._on_cell_selected)
        self.formula_bar.returnPressed.connect(self._apply_formula_from_bar)

    # ==================================================
    # HEADERS
    # ==================================================
    def _setup_headers(self):
        # Columns: A, B, C...
        for c in range(COLS):
            self.table.setHorizontalHeaderItem(
                c, QTableWidgetItem(chr(ord("A") + c))
            )

        # Rows: 1, 2, 3...
        for r in range(ROWS):
            self.table.setVerticalHeaderItem(
                r, QTableWidgetItem(str(r + 1))
            )

    # ==================================================
    # TABLE BEHAVIOR
    # ==================================================
    def _setup_table_behavior(self):
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)

        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked |
            QAbstractItemView.EditKeyPressed |
            QAbstractItemView.AnyKeyPressed
        )

        self.table.setShowGrid(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(24)

    # ==================================================
    # FORMULA BAR LOGIC
    # ==================================================
    def _on_cell_selected(self, current, previous):
        if not current:
            self.formula_bar.clear()
            return

        formula = current.data(Qt.UserRole)
        if formula:
            self.formula_bar.setText("=" + formula)
        else:
            self.formula_bar.setText(current.text())

    def _apply_formula_from_bar(self):
        item = self.table.currentItem()
        if not item:
            return

        text = self.formula_bar.text().strip()

        # signal loop'u kır
        self.table.blockSignals(True)
        item.setText(text)
        self.table.blockSignals(False)

        # engine manuel çağır
        self.engine.process_item(item)