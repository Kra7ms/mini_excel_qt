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
    QComboBox,
    QColorDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication, QColor, QBrush
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

        self.border_button = QToolButton()
        self.border_button.setText("▦")
        self.border_button.setPopupMode(QToolButton.InstantPopup)
        self.border_button.setFixedWidth(32)
        bar.addWidget(self.border_button)
        self._setup_border_menu()

        self.fill_button = QPushButton("🪣")
        self.fill_button.setFixedWidth(32)
        bar.addWidget(self.fill_button)

        self.text_color_button = QPushButton("🅰️")
        self.text_color_button.setFixedWidth(32)
        bar.addWidget(self.text_color_button)

        self.font_box = QFontComboBox()
        self.font_box.setMaximumWidth(160)
        bar.addWidget(self.font_box)

        self.font_size_box = QComboBox()
        self.font_size_box.setEditable(True)
        self.font_size_box.setMaximumWidth(60)

        self.bold_button = QToolButton()
        self.bold_button.setText("B")
        self.bold_button.setCheckable(True)
        self.bold_button.setFixedWidth(28)
        self.bold_button.setStyleSheet("font-weight: bold;")
        bar.addWidget(self.bold_button)

        self.align_left_btn = QToolButton()
        self.align_left_btn.setText("⟸")
        self.align_left_btn.setCheckable(True)
        self.align_left_btn.setFixedWidth(28)
        bar.addWidget(self.align_left_btn)

        self.align_center_btn = QToolButton()
        self.align_center_btn.setText("⟺")
        self.align_center_btn.setCheckable(True)
        self.align_center_btn.setFixedWidth(28)
        bar.addWidget(self.align_center_btn)

        self.align_right_btn = QToolButton()
        self.align_right_btn.setText("⟹")
        self.align_right_btn.setCheckable(True)
        self.align_right_btn.setFixedWidth(28)
        bar.addWidget(self.align_right_btn)

        self.wrap_button = QToolButton()
        self.wrap_button.setText("↩︎")
        self.wrap_button.setCheckable(True)
        self.wrap_button.setFixedWidth(28)
        bar.addWidget(self.wrap_button)

        self.merge_button = QToolButton()
        self.merge_button.setText("Merge")
        self.merge_button.setFixedWidth(48)
        bar.addWidget(self.merge_button)

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
        self.bold_button.clicked.connect(self._toggle_bold)
        self.fill_button.clicked.connect(self._choose_fill_color)
        self.text_color_button.clicked.connect(self._choose_text_color)
        self.align_left_btn.clicked.connect(lambda: self._set_alignment(Qt.AlignLeft))
        self.align_center_btn.clicked.connect(lambda: self._set_alignment(Qt.AlignCenter))
        self.align_right_btn.clicked.connect(lambda: self._set_alignment(Qt.AlignRight))
        self.wrap_button.clicked.connect(self._toggle_wrap)
        self.merge_button.clicked.connect(self._toggle_merge)

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
            (
                item.row(),
                item.column(),
                item.text(),
                item.data(Qt.UserRole),
                item.background(),
                item.foreground()
            )
        )


    def _undo(self):
        if not self.undo_stack:
            return
        row, col, text, formula, border, bg, fg = self.undo_stack.pop()
        item = self.table.item(row, col)
        if not item:
            return
        self._undo_block = True
        item.setText(text or "")
        item.setData(Qt.UserRole + 1, border)
        item.setData(Qt.UserRole, formula)
        item.setBackground(bg)
        item.setForeground(fg)
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

        self.bold_button.blockSignals(True)
        self.bold_button.setChecked(font.bold())
        self.bold_button.blockSignals(False)
        self._sync_alignment_buttons(current)

        align = current.textAlignment()
        self.wrap_button.blockSignals(True)
        self.wrap_button.setChecked(bool(align & Qt.TextWordWrap))
        self.wrap_button.blockSignals(False)

        span = self.table.span(current.row(), current.column())
        self.merge_button.setText("Unmerge" if span != (1, 1) else "Merge")

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
    # BUTONLAR
    # ==================================================
    def _toggle_bold(self):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        font = item.font()
        is_bold = font.bold()

        font.setBold(not is_bold)
        item.setFont(font)

    def _setup_border_menu(self):
        menu = QMenu(self)

        menu.addAction("No Border", lambda: self._set_border(None))
        menu.addAction("All Borders", lambda: self._set_border("all"))

        self.border_button.setMenu(menu)

    def _set_border(self, mode):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        if mode is None:
            item.setData(Qt.UserRole + 1, None)
        else:
            item.setData(Qt.UserRole + 1, mode)

        self._apply_table_borders()

    def _choose_fill_color(self):
        item = self.table.currentItem()
        if not item:
            return

        color = QColorDialog.getColor(parent=self, title="Fill Color")
        if not color.isValid():
            return

        # Undo için eski state
        self._push_undo_state(item)

        item.setBackground(color)
    
    def _choose_text_color(self):
        item = self.table.currentItem()
        if not item:
            return

        color = QColorDialog.getColor(parent=self, title="Text Color")
        if not color.isValid():
            return

        self._push_undo_state(item)

        item.setForeground(QBrush(color))

    def _set_alignment(self, align_flag):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        # dikey ortalama + yatay hizalama
        item.setTextAlignment(align_flag | Qt.AlignVCenter)

        # butonları senkronla
        self._sync_alignment_buttons(item)

    def _sync_alignment_buttons(self, item):
        align = item.textAlignment()

        self.align_left_btn.blockSignals(True)
        self.align_center_btn.blockSignals(True)
        self.align_right_btn.blockSignals(True)

        self.align_left_btn.setChecked(bool(align & Qt.AlignLeft))
        self.align_center_btn.setChecked(bool(align & Qt.AlignCenter))
        self.align_right_btn.setChecked(bool(align & Qt.AlignRight))

        self.align_left_btn.blockSignals(False)
        self.align_center_btn.blockSignals(False)
        self.align_right_btn.blockSignals(False)

    def _toggle_wrap(self):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        align = item.textAlignment()

        if align & Qt.TextWordWrap:
            # wrap kapat
            align &= ~Qt.TextWordWrap
            self.wrap_button.setChecked(False)
        else:
            # wrap aç
            align |= Qt.TextWordWrap
            self.wrap_button.setChecked(True)

        item.setTextAlignment(align)

        row = item.row()
        self.table.resizeRowToContents(row)

    def _toggle_merge(self):
        ranges = self.table.selectedRanges()
        if not ranges:
            return

        r = ranges[0]   # Excel gibi: sadece ilk seçimi al
        row = r.topRow()
        col = r.leftColumn()
        row_span = r.rowCount()
        col_span = r.columnCount()

        # Tek hücre → merge yapma
        if row_span == 1 and col_span == 1:
            return

        # Undo için (basit)
        self.undo_stack.append(
            ("merge", row, col, row_span, col_span)
        )

        # Eğer zaten merge ise → unmerge
        current_span = self.table.span(row, col)
        if current_span != (1, 1):
            self.table.setSpan(row, col, 1, 1)
            return

        # Merge
        self.table.setSpan(row, col, row_span, col_span)

    def _apply_table_borders(self):
        css = []

        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if not item:
                    continue

                border = item.data(Qt.UserRole + 1)
                if border == "all":
                    css.append(
                        f"QTableWidget::item(row:{r}, column:{c})"
                        "{ border: 1px solid black; }"
                    )

        self.table.setStyleSheet("\n".join(css))

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
