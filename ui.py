from PySide6.QtWidgets import (
    QMainWindow,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QLabel,
    QToolButton,
    QMenu,
    QWidgetAction,
    QListWidget
)
from PySide6.QtCore import Qt

from formula_engine import FormulaEngine
from utils import index_to_cell

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

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)

        # ===============================
        # FORMULA BAR ROW
        # ===============================
        bar_layout = QHBoxLayout()

        # Cell label (A1)
        self.cell_label = QLabel("A1")
        self.cell_label.setFixedWidth(40)
        self.cell_label.setAlignment(Qt.AlignCenter)
        bar_layout.addWidget(self.cell_label)

        # fx button
        self.fx_button = QToolButton()
        self.fx_button.setText("fx")
        self.fx_button.setPopupMode(QToolButton.InstantPopup)
        bar_layout.addWidget(self.fx_button)

        # Formula input
        self.formula_bar = QLineEdit()
        self.formula_bar.setPlaceholderText("Formula")
        bar_layout.addWidget(self.formula_bar)

        main_layout.addLayout(bar_layout)

        # ===============================
        # TABLE
        # ===============================
        self.table = QTableWidget(ROWS, COLS)
        main_layout.addWidget(self.table)

        self._setup_headers()
        self._setup_table_behavior()

        # ===============================
        # ENGINE
        # ===============================
        self.engine = FormulaEngine(self.table)

        # ===============================
        # FUNCTION MENU (fx)
        # ===============================
        self._setup_function_menu()

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
        for c in range(COLS):
            self.table.setHorizontalHeaderItem(
                c, QTableWidgetItem(chr(ord("A") + c))
            )
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
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(24)

    # ==================================================
    # FUNCTION MENU (QWidgetAction)
    # ==================================================
    def _setup_function_menu(self):
        menu = QMenu(self)

        action = QWidgetAction(menu)

        self.func_list = QListWidget()
        self.func_list.setMinimumWidth(180)
        self.func_list.setMaximumHeight(220)

        functions = [
            "SUM",
            "AVERAGE",
            "MIN",
            "MAX",
            "COUNT",
            "IF",
            "AND",
            "OR",
            "NOT",
        ]

        self.func_list.addItems(functions)

        action.setDefaultWidget(self.func_list)
        menu.addAction(action)

        self.func_list.itemClicked.connect(
            lambda item: self._insert_function(item.text())
        )

        self.fx_button.setMenu(menu)

    # ==================================================
    # FORMULA BAR LOGIC
    # ==================================================
    def _on_cell_selected(self, current, previous):
        if not current:
            return

        row, col = current.row(), current.column()
        self.cell_label.setText(index_to_cell(row, col))

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

        self.table.blockSignals(True)
        item.setText(text)
        self.table.blockSignals(False)

        self.engine.process_item(item)

    # ==================================================
    # INSERT FUNCTION
    # ==================================================
    def _insert_function(self, fn_name: str):
        cursor = self.formula_bar.cursorPosition()
        text = self.formula_bar.text()

        insert = f"{fn_name}()"
        self.formula_bar.setText(text[:cursor] + insert + text[cursor:])
        self.formula_bar.setCursorPosition(cursor + len(fn_name) + 1)
        self.formula_bar.setFocus()
