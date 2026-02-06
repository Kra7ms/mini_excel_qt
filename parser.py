import re
from typing import List

from ast_nodes import (
    ASTNode,
    Number,
    Cell,
    Range,
    BinaryOp,
    Function,
    Lambda,
    UnaryOp
)

# ======================================================
# TOKEN
# ======================================================

class Token:
    def __init__(self, type_, value):
        self.type = type_
        self.value = value

    def __repr__(self):
        return f"Token({self.type}, {self.value})"
    
# ======================================================
# TOKENIZER
# ======================================================

TOKEN_SPEC = [
    ("NUMBER",   r"\d+(\.\d+)?"),
    ("CELL",     r"[A-Z]+[0-9]+"),
    ("OP",       r"<=|>=|<>|==|!=|[+\-*/<>]=?"),
    ("COMMA",    r","),
    ("COLON",    r":"),
    ("LPAREN",   r"\("),
    ("RPAREN",   r"\)"),
    ("NAME",     r"[A-Z_]+"),
    ("SKIP",     r"[ \t]+"),
]

MASTER_RE = re.compile("|".join(
    f"(?P<{name}>{pattern})" for name, pattern in TOKEN_SPEC
))

def tokenize(text: str) -> List[Token]:
    tokens = []
    for match in MASTER_RE.finditer(text):
        kind = match.lastgroup
        value = match.group()

        if kind == "SKIP":
            continue

        tokens.append(Token(kind, value))

    return tokens

# ======================================================
# PARSER
# ======================================================

class Parser:
    def __init__(self):
        self.tokens = []
        self.pos = 0

    # ------------------
    # Helpers
    # ------------------

    def current(self):
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def eat(self, token_type=None):
        token = self.current()
        if token is None:
            return None

        if token_type and token.type != token_type:
            raise SyntaxError(f"Beklenen {token_type}, gelen {token.type}")

        self.pos += 1
        return token

    # ==================================================
    # ENTRY
    # ==================================================

    def parse(self, formula: str) -> ASTNode:
        self.tokens = tokenize(formula)
        self.pos = 0

        node = self.expression()
        if self.current() is not None:
            raise SyntaxError("Beklenmeyen token")
        return node

    # ==================================================
    # GRAMMAR
    # ==================================================

    # expression → comparison
    def expression(self):
        return self.comparison()

    # comparison → term ( (== | != | < | > | <= | >=) term )*
    def comparison(self):
        node = self.term()

        while self.current() and self.current().type == "OP":
            op = self.eat("OP").value
            right = self.term()
            node = BinaryOp(node, op, right)

        return node

    # term → factor ( (+ | -) factor )*
    def term(self):
        node = self.factor()

        while self.current() and self.current().value in ("+", "-"):
            op = self.eat("OP").value
            right = self.factor()
            node = BinaryOp(node, op, right)

        return node
    
    def power(self):
        node = self.unary()

        if self.current() and self.current().value == "^":
            self.eat("OP")
            right = self.power()   # sağa bağlı (right associative)
            return BinaryOp(node, "^", right)

        return node

    # factor → primary ( (* | /) primary )*
    def factor(self):
        node = self.power()

        while self.current() and self.current().value in ("*", "/"):
            op = self.eat("OP").value
            right = self.power()
            node = BinaryOp(node, op, right)

        return node
    
    def unary(self):
        token = self.current()

        if token and token.type == "OP" and token.value in ("+", "-"):
            op = self.eat("OP").value
            operand = self.unary()
            return UnaryOp(op, operand)

        return self.primary()

    # primary
    def primary(self):
        token = self.current()

        if token.type == "NUMBER":
            self.eat("NUMBER")
            return Number(float(token.value))

        if token.type == "CELL":
            start = self.eat("CELL").value

            if self.current() and self.current().type == "COLON":
                self.eat("COLON")
                end = self.eat("CELL").value
                return Range(Cell(start), Cell(end))

            return Cell(start)

        if token.type == "NAME":
            return self.function_or_name()

        if token.type == "LPAREN":
            self.eat("LPAREN")
            node = self.expression()
            self.eat("RPAREN")
            return node

        raise SyntaxError(f"Beklenmeyen token: {token}")
    
    # ==================================================
    # FUNCTIONS & LAMBDA
    # ==================================================

    def function_or_name(self):
        name = self.eat("NAME").value.upper()

        if self.current() and self.current().type == "LPAREN":
            self.eat("LPAREN")
            args = self.arguments()
            self.eat("RPAREN")

            if name == "LAMBDA":
                params = []
                for a in args[:-1]:
                    if isinstance(a, Cell):
                        params.append(a.ref)
                    else:
                        raise SyntaxError("LAMBDA parametreleri isim olmalı")

                return Lambda(params, args[-1])

            return Function(name, args)

        raise SyntaxError(f"Fonksiyon çağrısı bekleniyor: {name}")

    def arguments(self):
        args = []
        if self.current() and self.current().type != "RPAREN":
            args.append(self.expression())
            while self.current() and self.current().type == "COMMA":
                self.eat("COMMA")
                args.append(self.expression())
        return args