from typing import Any, Dict, List
from ast_nodes import (
    ASTNode,
    Number,
    Cell,
    Range,
    BinaryOp,
    Function,
    Lambda,
)
from utils import cell_to_index


class EvaluationError(Exception):
    pass


class Evaluator:
    def __init__(self, table):
        """
        table: QTableWidget veya cell(row, col) -> value sağlayan nesne
        """
        self.table = table

    # =====================================================
    # PUBLIC API
    # =====================================================
    def eval(self, node: ASTNode, env: Dict[str, Any] | None = None) -> Any:
        if env is None:
            env = {}

        if isinstance(node, Number):
            return node.value

        if isinstance(node, Cell):
            return self._eval_cell(node.ref)

        if isinstance(node, Range):
            return self._eval_range(node.start, node.end)

        if isinstance(node, BinaryOp):
            return self._eval_binary(node, env)

        if isinstance(node, Function):
            return self._eval_function(node, env)

        if isinstance(node, Lambda):
            return self._eval_lambda(node, env)

        raise EvaluationError(f"Bilinmeyen AST node: {node}")

    # =====================================================
    # CELL / RANGE
    # =====================================================
    def _eval_cell(self, ref: str) -> Any:
        idx = cell_to_index(ref)
        if not idx:
            return 0

        row, col = idx
        item = self.table.item(row, col)

        if not item or not item.text().strip():
            return 0

        text = item.text().strip()

        try:
            return float(text)
        except ValueError:
            return text

    def _eval_range(self, start: str, end: str) -> List[Any]:
        s = cell_to_index(start)
        e = cell_to_index(end)
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
    # BINARY OPERATORS
    # =====================================================
    def _eval_binary(self, node: BinaryOp, env: Dict[str, Any]):
        left = self.eval(node.left, env)
        right = self.eval(node.right, env)

        match node.op:
            case "+":
                return left + right
            case "-":
                return left - right
            case "*":
                return left * right
            case "/":
                return left / right
            case "^":
                return left ** right

            case "==":
                return left == right
            case "!=":
                return left != right
            case "<":
                return left < right
            case "<=":
                return left <= right
            case ">":
                return left > right
            case ">=":
                return left >= right

        raise EvaluationError(f"Bilinmeyen operator: {node.op}")
    
    def _flatten(self, args):
        for x in args:
            if isinstance(x, list):
                yield from x
            else:
                yield x

    # =====================================================
    # FUNCTIONS
    # =====================================================
    def _eval_function(self, node: Function, env: Dict[str, Any]):
        name = node.name.upper()
        args = [self.eval(arg, env) for arg in node.args]

        # -------- LOGIC --------
        if name == "IF":
            cond, a, b = args
            return a if bool(cond) else b

        if name == "AND":
            return all(bool(x) for x in args)

        if name == "OR":
            return any(bool(x) for x in args)

        if name == "NOT":
            return not bool(args[0])

        # -------- AGGREGATES --------
        if name == "SUM":
            return sum(self._flatten(args))

        if name == "AVERAGE":
            flat = self._flatten(args)
            return sum(flat) / len(flat) if flat else 0

        if name == "MIN":
            flat = self._flatten(args)
            return min(flat) if flat else 0

        if name == "MAX":
            flat = self._flatten(args)
            return max(flat) if flat else 0

        if name == "COUNT":
            flat = self._flatten(args)
            return len(flat)

        raise EvaluationError(f"Bilinmeyen fonksiyon: {name}")

    # =====================================================
    # LAMBDA
    # =====================================================
    def _eval_lambda(self, node: Lambda, env: Dict[str, Any]):
        def fn(*values):
            if len(values) != len(node.params):
                raise EvaluationError("LAMBDA argüman sayısı uyuşmuyor")

            local_env = env.copy()
            local_env.update(dict(zip(node.params, values)))

            return self.eval(node.body, local_env)

        return fn
