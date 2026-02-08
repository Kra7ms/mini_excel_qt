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
        self.clipboard_button.setPopupMode(QToolButton.InstantPopup)
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

        self.number_format_box = QComboBox()
        self.number_format_box.setMaximumWidth(140)
        self.number_format_box.addItems([
            "General",
            "Number (2 decimals)",
            "Integer",
            "Percent",
            "Currency (₺)",
        ])
        bar.addWidget(self.number_format_box)

        # Accounting button
        self.accounting_button = QToolButton()
        self.accounting_button.setText("₺≡")
        self.accounting_button.setFixedWidth(36)
        bar.addWidget(self.accounting_button)

        # Decrease Decimal
        self.dec_dec_button = QToolButton()
        self.dec_dec_button.setText("-.0")
        self.dec_dec_button.setFixedWidth(36)
        bar.addWidget(self.dec_dec_button)

        # Increase Decimal
        self.inc_dec_button = QToolButton()
        self.inc_dec_button.setText("+.0")
        self.inc_dec_button.setFixedWidth(36)
        bar.addWidget(self.inc_dec_button)

        self.autosum_button = QToolButton()
        self.autosum_button.setText("Σ")
        self.autosum_button.setFixedWidth(32)
        bar.addWidget(self.autosum_button)

        self.sort_asc_button = QToolButton()
        self.sort_asc_button.setText("A→Z")
        self.sort_asc_button.setFixedWidth(40)
        bar.addWidget(self.sort_asc_button)

        self.sort_desc_button = QToolButton()
        self.sort_desc_button.setText("Z→A")
        self.sort_desc_button.setFixedWidth(40)
        bar.addWidget(self.sort_desc_button)

        self.filter_button = QToolButton()
        self.filter_button.setText("Filter")
        self.filter_button.setCheckable(True)
        self.filter_button.setFixedWidth(50)
        bar.addWidget(self.filter_button)

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
        self.number_format_box.currentTextChanged.connect(self._change_number_format)
        self.accounting_button.clicked.connect(self._set_accounting)
        self.inc_dec_button.clicked.connect(self._increase_decimal)
        self.dec_dec_button.clicked.connect(self._decrease_decimal)
        self.autosum_button.clicked.connect(self._auto_sum)
        self.sort_asc_button.clicked.connect(lambda: self._sort_column(Qt.AscendingOrder))
        self.sort_desc_button.clicked.connect(lambda: self._sort_column(Qt.DescendingOrder))
        self.filter_button.clicked.connect(self._toggle_filter)

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
                item.data(Qt.UserRole + 1),
                item.background(),
                item.foreground(),
                item.data(Qt.UserRole + 2)
            )
        )
        


    def _undo(self):
        if not self.undo_stack:
            return
        row, col, text, formula, border, bg, fg, number_format = self.undo_stack.pop()
        item = self.table.item(row, col)
        if not item:
            return
        self._undo_block = True
        item.setText(text or "")
        item.setData(Qt.UserRole, formula)
        item.setData(Qt.UserRole + 1, border)
        item.setBackground(bg)
        item.setForeground(fg)
        item.setData(Qt.UserRole + 2, number_format)
        self._undo_block = False
        self.engine.process_item(item)

    def _on_item_changed(self, item):
        if self._undo_block:
            return
        self._push_undo_state(item)
        self.engine.process_item(item)
        self._apply_number_format(item)

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

        row = current.row()
        col = current.column()

        row_span = self.table.rowSpan(row, col)
        col_span = self.table.columnSpan(row, col)

        is_merged = (row_span > 1 or col_span > 1)

        fmt = current.data(Qt.UserRole + 2)
        self.number_format_box.blockSignals(True)
        self.number_format_box.setCurrentText(fmt if fmt else "General")
        self.number_format_box.blockSignals(False)


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

        item.setBackground(QBrush(color))
    
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

        # Eğer zaten merge ise → unmerge
        row_span_cur = self.table.rowSpan(row, col)
        col_span_cur = self.table.columnSpan(row, col)

        if row_span_cur > 1 or col_span_cur > 1:
            self.table.setSpan(row, col, 1, 1)
            return

        # Merge
        self.table.setSpan(row, col, row_span, col_span)

    def _change_number_format(self, fmt):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        if fmt == "General":
            item.setData(Qt.UserRole + 2, "General")

        elif fmt == "Integer":
            item.setData(Qt.UserRole + 2, "Integer:0")

        elif fmt == "Number (2 decimals)":
            item.setData(Qt.UserRole + 2, "Number:2")

        elif fmt == "Percent":
            item.setData(Qt.UserRole + 2, "Percent:0")

        elif fmt == "Currency (₺)":
            item.setData(Qt.UserRole + 2, "Currency:2")

        self._apply_number_format(item)

    def _apply_number_format(self, item):
            fmt = item.data(Qt.UserRole + 2)
            if not fmt:
                return

            raw = item.data(Qt.UserRole)
            if raw is None:
                return

            # raw string ise float'a çevirmeyi dene
            try:
                value = float(raw)
            except Exception:
                return
            
            if ":" in fmt:
                kind, dec = fmt.split(":")
                dec = int(dec)
            else:
                kind, dec = fmt, 0

            if kind == "General":
                text = str(raw)

            elif kind == "Integer":
                text = str(int(value))

            elif kind == "Number (2 decimals)":
                text = f"{value:.2f}"

            elif kind == "Percent":
                text = f"{value * 100:.0f}%"

            elif kind == "Currency (₺)":
                text = f"₺{value:,.2f}"

                        
            elif kind == "Accounting":
                text = f"₺ {value:,.{dec}f}"

            else:
                return

            self.table.blockSignals(True)
            item.setText(text)
            self.table.blockSignals(False)

    def _set_accounting(self):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        # varsayılan: 2 ondalık
        item.setData(Qt.UserRole + 2, "Accounting:2")
        self._apply_number_format(item)
    
    def _get_format_parts(self, item):
        fmt = item.data(Qt.UserRole + 2)
        if not fmt:
            return "General", 0

        if ":" in fmt:
            kind, dec = fmt.split(":")
            return kind, int(dec)

        return fmt, 0

    def _increase_decimal(self):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        kind, dec = self._get_format_parts(item)
        dec += 1

        item.setData(Qt.UserRole + 2, f"{kind}:{dec}")
        self._apply_number_format(item)

    def _decrease_decimal(self):
        item = self.table.currentItem()
        if not item:
            return

        self._push_undo_state(item)

        kind, dec = self._get_format_parts(item)
        dec = max(0, dec - 1)

        item.setData(Qt.UserRole + 2, f"{kind}:{dec}")
        self._apply_number_format(item)

    def _auto_sum(self):
        item = self.table.currentItem()
        if not item:
            return

        row = item.row()
        col = item.column()

        start = row - 1
        while start >= 0:
            cell = self.table.item(start, col)
            if not cell:
                break
            try:
                float(cell.text())
            except Exception:
                break
            start -= 1

        start += 1
        end = row - 1

        if start > end:
            return

        start_cell = index_to_cell(start, col)
        end_cell = index_to_cell(end, col)

        formula = f"=SUM({start_cell}:{end_cell})"

        self._push_undo_state(item)
        self.table.blockSignals(True)
        item.setText(formula)
        self.table.blockSignals(False)
        self.engine.process_item(item)

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
    
    def _sort_column(self, order):
        item = self.table.currentItem()
        if not item:
            return

        col = item.column()
        self.table.sortItems(col, order)

    def _toggle_filter(self):
        item = self.table.currentItem()
        if not item:
            return

        col = item.column()
        enabled = self.filter_button.isChecked()

        for r in range(self.table.rowCount()):
            cell = self.table.item(r, col)
            hide = enabled and (not cell or not cell.text())
            self.table.setRowHidden(r, hide)

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
