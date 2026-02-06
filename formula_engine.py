from PySide6.QtCore import Qt
from utils import index_to_cell
from parser import Parser
from ast_nodes import *
from dependency_graph import DependencyGraph
from dependency import DependencyExtractor
from evaluator import Evaluator

RANGE_FUNCTIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT")
LOGIC_FUNCTIONS = ("AND", "OR", "XOR")

class FormulaEngine:
    def __init__(self, table):
        self.table = table
        self.parser = Parser()
        self.evaluator = Evaluator(table)
        self.extractor = DependencyExtractor()
        self.graph = DependencyGraph()

    # =====================================================
    # ENTRY POINT
    # =====================================================
    def process_item(self, item):
        text = item.text().strip()
        row, col = item.row(), item.column()
        cell_ref = index_to_cell(row, col)


        if not text.startswith("="):
            item.setData(Qt.UserRole, None)
            self.graph.remove_cell(cell_ref)
            self._recalculate_dependents(cell_ref)
            return

        # ---------------------------
        # FORMÜL
        # ---------------------------
        formula = text[1:]
        item.setData(Qt.UserRole, formula)

        try:
            ast = self.parser.parse(formula)
        except Exception:
            self._set_item_value(item, "#PARSE!")
            return

        # Dependency çıkar
        deps = self.extractor.extract(ast)
        self.graph.set_dependencies((row, col), deps)

        # Evaluate
        try:
            value = self.evaluator.eval(ast)
        except Exception:
            value = "#ERROR"

        # UI güncelle
        self._set_item_value(item, value)

        # Bağımlıları güncelle
        self._recalculate_dependents(cell_ref)
    
    # =====================================================
    # DIŞTAN YENİDEN HESAPLAMA
    # =====================================================
    def _recalculate_dependents(self, cell_ref):
        for dep in self.graph.recalculate_dependents(cell_ref):
            r, c = self._cell_to_index(dep)
            self._recalculate_cell(r, c)

    def _recalculate_cell(self, row: int, col: int):
        item = self.table.item(row, col)
        if not item:
            return

        formula = item.data(Qt.UserRole)
        if not formula:
            return

        try:
            ast = self.parser.parse(formula)
            value = self.evaluator.eval(ast)
        except Exception:
            value = "#ERROR"

        self._set_item_value(item, value)

    # =====================================================
    # UI SAFE UPDATE
    # =====================================================
    def _set_item_value(self, item, value):
        table = self.table
        table.blockSignals(True)

        if value is True:
            item.setText("TRUE")
        elif value is False:
            item.setText("FALSE")
        else:
            item.setText(str(value))

        table.blockSignals(False)

    def _cell_to_index(self, ref):
        from utils import cell_to_index
        return cell_to_index(ref)