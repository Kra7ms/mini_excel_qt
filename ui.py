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
    QListWidget,
    QPushButton,
    QFontComboBox,
    QComboBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication

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
        # STATE
        # ===============================
        self.undo_stack = []
        self._undo_block = False
        self._format_painter_active = False
        self._copied_format = None

        # ===============================
        # CENTRAL + LAYOUT
        # ===============================
        central = QWidget(self)
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(4)

        # ===============================
        # TOP BAR
        # ===============================
        bar = QHBoxLayout()

        self.cell_label = QLabel("A1")
        self.cell_label.setFixedWidth(40)
        self.cell_label.setAlignment(Qt.AlignCenter)
        bar.addWidget(self.cell_label)

        self.fx_button = QToolButton()
        self.fx_button.setText("fx")
        self.fx_button.setPopupMode(QToolButton.InstantPopup)
        bar.addWidget(self.fx_button)

        self.undo_button = QPushButton("⟲")
        self.undo_button.setFixedWidth(32)
        bar.addWidget(self.undo_button)

        self.clipboard_button = QToolButton()
        self.clipboard_button.setText("📋")
        self.clipboard_button.setFixedWidth(32)
        bar.addWidget(self.clipboard_button)

        self.format_button = QPushButton("🖌️")
        self.format_button.setCheckable(True)
        self.format_button.setFixedWidth(32)
        bar.addWidget(self.format_button)

        self.font_box = QFontComboBox()
        self.font_box.setMaximumWidth(160)
        bar.addWidget(self.font_box)

        self.font_size_box = QComboBox()
        self.font_size_box.setEditable(True)
        self.font_size_box.setMaximumWidth(60)

        sizes = [
            "8", "9", "10", "11", "12", "14", "16",
            "18", "20", "22", "24", "26", "28", "36", "48", "72"
        ]
        self.font_size_box.addItems(sizes)
        self.font_size_box.setCurrentText("11")

        bar.addWidget(self.font_size_box)

        self.formula_bar = QLineEdit()
        self.formula_bar.setPlaceholderText("Formula")
        bar.addWidget(self.formula_bar)

        main_layout.addLayout(bar)

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
        # MENUS
        # ===============================
        self._setup_function_menu()
        self._setup_clipboard_menu()

        # ===============================
        # SIGNALS
        # ===============================
        self.table.itemChanged.connect(self._on_item_changed)
        self.table.currentItemChanged.connect(self._on_cell_selected)
        self.formula_bar.returnPressed.connect(self._apply_formula_from_bar)
        self.undo_button.clicked.connect(self._undo)
        self.format_button.clicked.connect(self._toggle_format_painter)
        self.font_box.currentFontChanged.connect(self._change_font)
        self.font_size_box.currentTextChanged.connect(self._change_font_size)

    # ==================================================
    # SETUP
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
    # FUNCTION MENU
    # ==================================================
    def _setup_function_menu(self):
        menu = QMenu(self)
        action = QWidgetAction(menu)

        self.func_list = QListWidget()
        self.func_list.setMinimumWidth(180)
        self.func_list.setMaximumHeight(220)

        self.func_list.addItems([
            "SUM", "AVERAGE", "MIN", "MAX", "COUNT",
            "IF", "AND", "OR", "NOT"
        ])

        action.setDefaultWidget(self.func_list)
        menu.addAction(action)

        self.func_list.itemClicked.connect(
            lambda item: self._insert_function(item.text())
        )

        self.fx_button.setMenu(menu)

    # ==================================================
    # CLIPBOARD
    # ==================================================
    def _setup_clipboard_menu(self):
        menu = QMenu(self)
        menu.addAction("Copy", self._copy_cell)
        menu.addAction("Cut", self._cut_cell)
        menu.addAction("Paste", self._paste_cell)

        self.clipboard_button.setMenu(menu)


    def _copy_cell(self):
        item = self.table.currentItem()
        if item:
            QGuiApplication.clipboard().setText(item.text())

    def _cut_cell(self):
        item = self.table.currentItem()
        if not item:
            return
        self._push_undo_state(item)
        QGuiApplication.clipboard().setText(item.text())
        self.table.blockSignals(True)
        item.setText("")
        item.setData(Qt.UserRole, None)
        self.table.blockSignals(False)
        self.engine.process_item(item)

    def _paste_cell(self):
        item = self.table.currentItem()
        if not item:
            return
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        self._push_undo_state(item)
        self.table.blockSignals(True)
        item.setText(text)
        self.table.blockSignals(False)
        self.engine.process_item(item)

    # ==================================================
    # UNDO
    # ==================================================
    def _push_undo_state(self, item):
        if self._undo_block:
            return
        self.undo_stack.append(
            (item.row(), item.column(), item.text(), item.data(Qt.UserRole))
        )

    def _undo(self):
        if not self.undo_stack:
            return
        row, col, text, formula = self.undo_stack.pop()
        item = self.table.item(row, col)
        if not item:
            return
        self._undo_block = True
        item.setText(text or "")
        item.setData(Qt.UserRole, formula)
        self._undo_block = False
        self.engine.process_item(item)

    def _on_item_changed(self, item):
        if self._undo_block:
            return
        self._push_undo_state(item)
        self.engine.process_item(item)

    # ==================================================
    # FORMAT PAINTER
    # ==================================================
    def _toggle_format_painter(self):
        item = self.table.currentItem()
        if not item:
            self.format_button.setChecked(False)
            return
        if self.format_button.isChecked():
            self._copied_format = {
                "font": item.font(),
                "foreground": item.foreground(),
                "background": item.background(),
                "alignment": item.textAlignment(),
            }
            self._format_painter_active = True
        else:
            self._format_painter_active = False

    def _apply_format(self, item):
        f = self._copied_format
        if not f:
            return
        item.setFont(f["font"])
        item.setForeground(f["foreground"])
        item.setBackground(f["background"])
        item.setTextAlignment(f["alignment"])

    # ==================================================
    # FONT
    # ==================================================
    def _change_font(self, font):
        item = self.table.currentItem()
        if not item:
            return
        self._push_undo_state(item)
        f = item.font()
        f.setFamily(font.family())
        item.setFont(f)

    def _change_font_size(self, size_text):
        item = self.table.currentItem()
        if not item:
            return

        try:
            size = int(size_text)
        except ValueError:
            return

        self._push_undo_state(item)

        font = item.font()
        font.setPointSize(size)
        item.setFont(font)


    # ==================================================
    # FORMULA BAR
    # ==================================================
    def _on_cell_selected(self, current, previous):
        if self._format_painter_active and current:
            self._apply_format(current)
            self._format_painter_active = False
            self.format_button.setChecked(False)

        if not current:
            return

        self.cell_label.setText(index_to_cell(current.row(), current.column()))

        formula = current.data(Qt.UserRole)
        self.formula_bar.setText(
            "=" + formula if formula else current.text()
        )

        self.font_box.blockSignals(True)
        self.font_box.setCurrentFont(current.font())
        self.font_box.blockSignals(False)

        font = current.font()
        self.font_size_box.blockSignals(True)
        self.font_size_box.setCurrentText(str(font.pointSize()))
        self.font_size_box.blockSignals(False)

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
    def _insert_function(self, fn):
        cursor = self.formula_bar.cursorPosition()
        text = self.formula_bar.text()
        insert = f"{fn}()"
        self.formula_bar.setText(text[:cursor] + insert + text[cursor:])
        self.formula_bar.setCursorPosition(cursor + len(fn) + 1)
        self.formula_bar.setFocus()
