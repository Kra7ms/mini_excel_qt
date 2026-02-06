from ast_nodes import (
    ASTNode, Cell, Range,
    BinaryOp, Function, Lambda, Number
)
from utils import expand_range

class DependencyExtractor:
    def extract(self, node: ASTNode) -> set:
        deps = set()
        self._walk(node, deps)
        return deps

    def _walk(self, node: ASTNode, deps: set):
        if isinstance(node, Cell):
            deps.add(node.ref)

        elif isinstance(node, Range):
            for cell in expand_range(node.start.ref, node.end.ref):
                deps.add(cell)

        elif isinstance(node, BinaryOp):
            self._walk(node.left, deps)
            self._walk(node.right, deps)

        elif isinstance(node, Function):
            for arg in node.args:
                self._walk(arg, deps)

        elif isinstance(node, Lambda):
            self._walk(node.body, deps)

        elif isinstance(node, Number):
            pass