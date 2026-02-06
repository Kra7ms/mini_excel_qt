from dataclasses import dataclass
from typing import List

# =========================
# BASE NODE
# =========================

class ASTNode:
    pass


# =========================
# LITERALS
# =========================

@dataclass(frozen=True)
class Number(ASTNode):
    value: float


@dataclass(frozen=True)
class String(ASTNode):
    value: str


@dataclass(frozen=True)
class Boolean(ASTNode):
    value: bool


# =========================
# CELL REFERENCES
# =========================

@dataclass(frozen=True)
class Cell(ASTNode):
    ref: str   # "A1"


@dataclass(frozen=True)
class Range(ASTNode):
    start: Cell
    end: Cell


# =========================
# SYMBOL (LAMBDA param)
# =========================

@dataclass(frozen=True)
class Symbol(ASTNode):
    name: str


# =========================
# OPERATIONS
# =========================

@dataclass(frozen=True)
class UnaryOp(ASTNode):
    op: str            # "-", "NOT"
    operand: ASTNode


@dataclass(frozen=True)
class BinaryOp(ASTNode):
    left: ASTNode
    op: str            # "+", "-", "*", "/", "<", "==", etc.
    right: ASTNode


# =========================
# FUNCTIONS
# =========================

@dataclass(frozen=True)
class Function(ASTNode):
    name: str
    args: List[ASTNode]


@dataclass(frozen=True)
class If(ASTNode):
    condition: ASTNode
    true_expr: ASTNode
    false_expr: ASTNode


@dataclass(frozen=True)
class Lambda(ASTNode):
    params: List[str]
    body: ASTNode