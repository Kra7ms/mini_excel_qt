import re
from PySide6.QtCore import Qt
from utils import cell_to_index

RANGE_FUNCTIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT", "SCAN")
LOGIC_FUNCTIONS = ("AND", "OR", "XOR")


class FormulaEngine:
    def __init__(self, table):
        self.table = table
        self.dependencies = {}
        self._updating = False

    # =====================================================
    # ANA GİRİŞ
    # =====================================================
    def process_item(self, item):
        if self._updating:
            return

        text = item.text().strip()
        row, col = item.row(), item.column()

        if not text.startswith("="):
            item.setData(Qt.UserRole, None)
            return

        self._updating = True
        item.setData(Qt.UserRole, text)

        value = self.evaluate(text)
        item.setText(value)

        self._register_dependencies(row, col, text)
        self._updating = False

    # =====================================================
    # FORMÜL DEĞERLENDİRME
    # =====================================================
    def evaluate(self, formula: str) -> str:
        expr = formula[1:].strip()

        # TRUE / FALSE sabitleri
        expr = re.sub(r"\bTRUE\(\)", "True", expr, flags=re.IGNORECASE)
        expr = re.sub(r"\bFALSE\(\)", "False", expr, flags=re.IGNORECASE)        

        expr = self._handle_if(expr)
        expr = self._handle_switch(expr)
        expr = self._handle_scan(expr)
        expr = self._handle_logic(expr)
        expr = self._handle_range_functions(expr)
        expr = self._replace_cells(expr)

        try:
            result = eval(expr, {"__builtins__": None}, {})
            return "TRUE" if result is True else "FALSE" if result is False else str(result)
        except:
            return "#ERROR"
        
    # =====================================================
    # IF FONKSİYONU
    # =====================================================
    def _handle_if(self, expr: str):
        if not expr.upper().startswith("IF("):
            return expr

        inside = expr[3:-1]
        args = self._split_args(inside)

        if len(args) != 3:
            return "#ERROR"

        condition, true_part, false_part = args

        try:
            cond = self.evaluate("=" + condition)
            result = bool(eval(cond))
        except Exception:
            result = False

        chosen = true_part if result else false_part
        return self.evaluate("=" + chosen.strip())
    
    # ====================================================
    # SWITCH FONKSİYONU
    # ====================================================
    def _handle_switch(self, expr: str):
        if not expr.upper().startswith("SWITCH("):
            return expr

        inside = expr[7:-1]  # SWITCH(...)
        args = self._split_if_args(inside)

        if len(args) < 3:
            return "#ERROR"

        # İlk ifade
        try:
            base_value = self.evaluate("=" + args[0])
            base_value = eval(base_value)
        except:
            return "#ERROR"

        pairs = args[1:]

        default = None
        if len(pairs) % 2 == 1:
            default = pairs[-1]
            pairs = pairs[:-1]

        for i in range(0, len(pairs), 2):
            try:
                match_val = eval(self.evaluate("=" + pairs[i]))
                if base_value == match_val:
                    return self.evaluate("=" + pairs[i + 1])
            except:
                continue

        if default is not None:
            return self.evaluate("=" + default)

        return ""
    
    # ====================================================
    # SCAN FONKSİYONU
    # ====================================================

    def _handle_scan(self, expr: str):
        pattern = r"SCAN\((.+?),(.+?),(.+?)\)"
        match = re.search(pattern, expr, re.IGNORECASE)

        if not match:
            return expr

        init_expr, range_expr, lambda_expr = match.groups()

        # Initial value
        try:
            acc = float(self.evaluate("=" + init_expr))
        except:
            return "#ERROR"

        # Range
        if ":" not in range_expr:
            return "#ERROR"

        start, end = range_expr.split(":")
        values = self._get_range_values(start.strip(), end.strip())

        results = []

        for v in values:
            try:
                expr_eval = lambda_expr.replace("a", str(acc)).replace("b", str(v))
                acc = eval(expr_eval, {"__builtins__": None}, {})
                results.append(acc)
            except:
                return "#ERROR"

        # Tek hücreye yazmak için
        return ",".join(str(r) for r in results)

    # =====================================================
    # AND / OR / XOR
    # =====================================================
    def _handle_logic(self, expr: str):
        for func in LOGIC_FUNCTIONS:
            pattern = rf"{func}\((.*?)\)"
            expr = re.sub(
                pattern,
                lambda m: self._eval_logic(m.group(1), func),
                expr,
                flags=re.IGNORECASE
            )
        return expr

    def _eval_logic(self, content: str, mode: str):
        parts = self._split_args(content)
        results = []

        for part in parts:
            try:
                val = self.evaluate("=" + part)
                results.append(bool(eval(val)))
            except:
                results.append(False)

        if mode == "AND":
            return str(all(results))
        if mode == "OR":
            return str(any(results))
        if mode == "XOR":
            return str(sum(results) == 1)

        return "False"

    # =====================================================
    # RANGE FONKSİYONLARI
    # =====================================================
    def _handle_range_functions(self, expr: str):
        for func in RANGE_FUNCTIONS:
            pattern = rf"{func}\(([A-Z]+[0-9]+):([A-Z]+[0-9]+)\)"
            for start, end in re.findall(pattern, expr):
                values = self._get_range_values(start, end)

                if func == "SUM":
                    result = sum(values)
                elif func == "AVERAGE":
                    result = sum(values) / len(values) if values else 0
                elif func == "MIN":
                    result = min(values) if values else 0
                elif func == "MAX":
                    result = max(values) if values else 0
                elif func == "COUNT":
                    result = len(values)

                expr = expr.replace(f"{func}({start}:{end})", str(result))

        return expr

    # =====================================================
    # HÜCRE DEĞİŞTİRME
    # =====================================================
    def _replace_cells(self, expr: str):
        expr = expr.replace("<>", "!=")
        expr = re.sub(r"(?<![<>=!])=(?!=)", "==", expr)

        refs = set(re.findall(r"[A-Z]+[0-9]+", expr))
        for ref in refs:
            expr = expr.replace(ref, self._get_cell_value(ref))

        return expr

    def _get_cell_value(self, ref: str):
        idx = cell_to_index(ref)
        if not idx:
            return "0"

        row, col = idx
        item = self.table.item(row, col)

        if not item or not item.text().strip():
            return "0"

        text = item.text().strip()
        try:
            float(text)
            return text
        except:
            return f'"{text}"'

    # =====================================================
    # RANGE DEĞERLERİ
    # =====================================================
    def _get_range_values(self, start, end):
        s, e = cell_to_index(start), cell_to_index(end)
        if not s or not e:
            return []

        r1, c1 = s
        r2, c2 = e

        values = []
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                item = self.table.item(r, c)
                if not item:
                    continue
                try:
                    values.append(float(item.text()))
                except:
                    pass
        return values

    # =====================================================
    # ARG SPLITTER (IF / AND / OR / XOR için)
    # =====================================================
    def _split_args(self, text):
        args, depth, current = [], 0, ""
        for ch in text:
            if ch == "," and depth == 0:
                args.append(current.strip())
                current = ""
            else:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                current += ch
        args.append(current.strip())
        return args

    # =====================================================
    # BAĞIMLILIK
    # =====================================================
    def _register_dependencies(self, row, col, formula):
        refs = set(re.findall(r"[A-Z]+[0-9]+", formula))
        self.dependencies[(row, col)] = {
            cell_to_index(r) for r in refs if cell_to_index(r)
        }

    def recalculate_all(self):
        for (r, c) in self.dependencies:
            item = self.table.item(r, c)
            if item:
                formula = item.data(Qt.UserRole)
                if formula:
                    item.setText(self.evaluate(formula))
