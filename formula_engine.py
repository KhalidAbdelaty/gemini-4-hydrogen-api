"""A deliberately small Excel formula engine for the Northstar workbook.

It parses formulas with openpyxl's Tokenizer and evaluates only this subset:
numbers, text, TRUE/FALSE, cell and range references (same-sheet or
Sheet!A1, with or without $), the operators + - * / & = <> < > <= >= and
unary minus, and the functions SUM, SUMIFS, IF, MIN, MAX, ROUND.
Anything else raises UnsupportedFormula. It is not Excel.
"""
from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

SUPPORTED_FUNCTIONS = {"SUM", "SUMIFS", "IF", "MIN", "MAX", "ROUND"}
_REF = re.compile(r"^(?:(?P<sheet>'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!)?(?P<a>\$?[A-Z]{1,3}\$?\d+)(?::(?P<b>\$?[A-Z]{1,3}\$?\d+))?$")


class UnsupportedFormula(ValueError):
    """The formula uses syntax or a function outside the supported subset."""


class FormulaError(ValueError):
    """Evaluation failed: circular reference, bad operand, unknown sheet."""


@dataclass(frozen=True)
class Ref:
    sheet: str
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    def cells(self):
        for r in range(self.min_row, self.max_row + 1):
            for c in range(self.min_col, self.max_col + 1):
                yield (self.sheet, f"{get_column_letter(c)}{r}")

    @property
    def is_single(self):
        return self.min_col == self.max_col and self.min_row == self.max_row

    def label(self):
        a = f"{get_column_letter(self.min_col)}{self.min_row}"
        b = f"{get_column_letter(self.max_col)}{self.max_row}"
        return f"{self.sheet}!{a}" if self.is_single else f"{self.sheet}!{a}:{b}"


def parse_ref(text: str, current_sheet: str) -> Ref:
    m = _REF.match(text.strip())
    if not m:
        raise UnsupportedFormula(f"unsupported reference: {text}")
    sheet = (m.group("sheet") or current_sheet).strip("'")
    area = m.group("a").replace("$", "")
    if m.group("b"):
        area += ":" + m.group("b").replace("$", "")
    min_col, min_row, max_col, max_row = range_boundaries(area)
    return Ref(sheet, min_col, min_row, max_col, max_row)


# ---------- parser: tokens -> small AST of tuples ----------
_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4}


def parse(formula: str, sheet: str):
    if not isinstance(formula, str) or not formula.startswith("="):
        raise UnsupportedFormula("not a formula")
    try:
        toks = [t for t in Tokenizer(formula).items if t.type != Token.WSPACE]
    except Exception as exc:  # tokenizer rejects malformed text
        raise UnsupportedFormula(f"cannot tokenize: {exc}") from exc
    pos = 0

    def peek():
        return toks[pos] if pos < len(toks) else None

    def take():
        nonlocal pos
        tok = toks[pos]
        pos += 1
        return tok

    def primary():
        tok = peek()
        if tok is None:
            raise UnsupportedFormula("unexpected end of formula")
        if tok.type == Token.OP_PRE and tok.value in "+-":
            take()
            return ("neg", primary()) if tok.value == "-" else primary()
        if tok.type == Token.OPERAND:
            take()
            if tok.subtype == Token.NUMBER:
                return ("num", float(tok.value))
            if tok.subtype == Token.TEXT:
                return ("str", tok.value[1:-1].replace('""', '"'))
            if tok.subtype == Token.LOGICAL:
                return ("bool", tok.value.upper() == "TRUE")
            if tok.subtype == Token.RANGE:
                return ("ref", parse_ref(tok.value, sheet))
            raise UnsupportedFormula(f"unsupported operand {tok.value}")
        if tok.type == Token.FUNC and tok.subtype == Token.OPEN:
            take()
            name = tok.value[:-1].upper()
            if name not in SUPPORTED_FUNCTIONS:
                raise UnsupportedFormula(f"unsupported function {name}")
            args = []
            if not (peek() and peek().type == Token.FUNC and peek().subtype == Token.CLOSE):
                while True:
                    args.append(expr(0))
                    nxt = take()
                    if nxt.type == Token.SEP and nxt.subtype == Token.ARG:
                        continue
                    if nxt.type == Token.FUNC and nxt.subtype == Token.CLOSE:
                        break
                    raise UnsupportedFormula(f"unexpected token {nxt.value}")
            else:
                take()
            return ("call", name, args)
        if tok.type == Token.PAREN and tok.subtype == Token.OPEN:
            take()
            node = expr(0)
            close = take()
            if close.type != Token.PAREN:
                raise UnsupportedFormula("unbalanced parentheses")
            return node
        raise UnsupportedFormula(f"unsupported token {tok.value!r}")

    def expr(min_prec):
        left = primary()
        while True:
            tok = peek()
            if tok is None or tok.type != Token.OP_IN or tok.value not in _PREC:
                if tok is not None and tok.type == Token.OP_IN:
                    raise UnsupportedFormula(f"unsupported operator {tok.value}")
                return left
            prec = _PREC[tok.value]
            if prec < min_prec:
                return left
            take()
            right = expr(prec + 1)
            left = ("bin", tok.value, left, right)

    tree = expr(0)
    if pos != len(toks):
        raise UnsupportedFormula(f"trailing tokens after {toks[pos - 1].value}")
    return tree


def references(formula: str, sheet: str) -> list[Ref]:
    """Every reference in a supported formula, in order of appearance."""
    found = []

    def walk(node):
        kind = node[0]
        if kind == "ref":
            found.append(node[1])
        elif kind == "neg":
            walk(node[1])
        elif kind == "bin":
            walk(node[2])
            walk(node[3])
        elif kind == "call":
            for arg in node[2]:
                walk(arg)

    walk(parse(formula, sheet))
    return found


def numeric_literals(formula: str, sheet: str) -> list[float]:
    out = []

    def walk(node):
        if node[0] == "num":
            out.append(node[1])
        elif node[0] == "neg":
            walk(node[1])
        elif node[0] == "bin":
            walk(node[2])
            walk(node[3])
        elif node[0] == "call":
            for arg in node[2]:
                walk(arg)

    walk(parse(formula, sheet))
    return out


# ---------- evaluation ----------
class Workbook:
    """Cells as {(sheet, 'A1'): value-or-formula}, evaluated on demand."""

    def __init__(self, cells: dict, sheets: list[str]):
        self.cells = dict(cells)
        self.sheets = list(sheets)
        self._cache: dict = {}
        self._stack: set = set()

    @classmethod
    def from_xlsx(cls, path) -> "Workbook":
        wb = load_workbook(path, data_only=False)
        cells = {}
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        cells[(ws.title, c.coordinate)] = c.value
        return cls(cells, wb.sheetnames)

    def copy(self) -> "Workbook":
        return Workbook(self.cells, self.sheets)

    def set(self, sheet, coord, value):
        self.cells[(sheet, coord)] = value
        self._cache.clear()

    def raw(self, sheet, coord):
        return self.cells.get((sheet, coord))

    def value(self, sheet, coord):
        key = (sheet, coord)
        if sheet not in self.sheets:
            raise FormulaError(f"unknown sheet {sheet}")
        if key in self._cache:
            return self._cache[key]
        raw = self.cells.get(key)
        if isinstance(raw, str) and raw.startswith("="):
            if key in self._stack:
                raise FormulaError(f"circular reference at {sheet}!{coord}")
            self._stack.add(key)
            try:
                result = self._eval(parse(raw, sheet))
            finally:
                self._stack.discard(key)
        else:
            result = raw
        self._cache[key] = result
        return result

    def formula_cells(self):
        return [k for k, v in self.cells.items() if isinstance(v, str) and v.startswith("=")]

    # -- helpers --
    def _range_values(self, ref: Ref):
        return [self.value(s, c) for s, c in ref.cells()]

    def _range_rows(self, ref: Ref):
        return list(ref.cells())

    @staticmethod
    def _num(v):
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        raise FormulaError(f"expected a number, got {v!r}")

    def _scalar(self, node):
        if node[0] == "ref":
            if not node[1].is_single:
                raise FormulaError(f"range {node[1].label()} used where one value is expected")
            s, c = next(node[1].cells())
            return self.value(s, c)
        return self._eval(node)

    def _eval(self, node):
        kind = node[0]
        if kind in ("num", "str", "bool"):
            return node[1]
        if kind == "ref":
            return self._scalar(node)
        if kind == "neg":
            return -self._num(self._eval(node[1]))
        if kind == "bin":
            op, a, b = node[1], self._eval(node[2]), self._eval(node[3])
            if op == "&":
                return f"{'' if a is None else a}{'' if b is None else b}"
            if op in ("+", "-", "*", "/"):
                x, y = self._num(a), self._num(b)
                if op == "/" and y == 0:
                    raise FormulaError("division by zero")
                return {"+": x + y, "-": x - y, "*": x * y, "/": x / y if y else 0}[op]
            return _compare(op, a, b)
        if kind == "call":
            return self._call(node[1], node[2])
        raise UnsupportedFormula(f"unknown node {kind}")

    def _flat_numbers(self, args):
        out = []
        for arg in args:
            if arg[0] == "ref":
                out += [float(v) for v in self._range_values(arg[1]) if isinstance(v, (int, float)) and not isinstance(v, bool)]
            else:
                out.append(self._num(self._eval(arg)))
        return out

    def _call(self, name, args):
        if name == "SUM":
            return sum(self._flat_numbers(args))
        if name in ("MIN", "MAX"):
            nums = self._flat_numbers(args)
            return (min if name == "MIN" else max)(nums) if nums else 0.0
        if name == "ROUND":
            if len(args) != 2:
                raise FormulaError("ROUND needs 2 arguments")
            return round(self._num(self._eval(args[0])), int(self._num(self._eval(args[1]))))
        if name == "IF":
            if len(args) not in (2, 3):
                raise FormulaError("IF needs 2 or 3 arguments")
            cond = self._eval(args[0])
            truthy = bool(cond) if not isinstance(cond, str) else cond != ""
            if truthy:
                return self._eval(args[1])
            return self._eval(args[2]) if len(args) == 3 else False
        if name == "SUMIFS":
            if len(args) < 3 or len(args) % 2 == 0:
                raise FormulaError("SUMIFS needs a sum range and criteria pairs")
            if any(a[0] != "ref" for a in (args[0], *args[1::2])):
                raise FormulaError("SUMIFS ranges must be references")
            sum_cells = self._range_rows(args[0][1])
            pairs = []
            for rng, crit in zip(args[1::2], args[2::2]):
                cells = self._range_rows(rng[1])
                if len(cells) != len(sum_cells):
                    raise FormulaError("SUMIFS ranges have different sizes")
                pairs.append((cells, self._eval(crit)))
            total = 0.0
            for i, (s, c) in enumerate(sum_cells):
                if all(_criterion(self.value(*cells[i]), crit) for cells, crit in pairs):
                    v = self.value(s, c)
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        total += float(v)
            return total
        raise UnsupportedFormula(f"unsupported function {name}")


def _compare(op, a, b):
    if isinstance(a, str) or isinstance(b, str):
        a, b = str(a or "").lower(), str(b or "").lower()
    else:
        a, b = Workbook._num(a), Workbook._num(b)
    return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]


def _criterion(value, crit):
    if isinstance(crit, str):
        m = re.match(r"^(<=|>=|<>|<|>|=)(.*)$", crit)
        if m:
            op, rhs = m.groups()
            try:
                return _compare(op, value, float(rhs))
            except ValueError:
                return _compare(op, value, rhs)
        return str(value or "").strip().lower() == crit.strip().lower()
    if value is None:
        return False
    try:
        return float(value) == float(crit)
    except (TypeError, ValueError):
        return False


def find_soffice() -> str:
    """LibreOffice's command-line binary: $SOFFICE, PATH, or the default Windows/macOS install."""
    candidates = [os.environ.get("SOFFICE"), shutil.which("soffice"), shutil.which("libreoffice"),
                  r"C:\Program Files\LibreOffice\program\soffice.exe",
                  "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    for c in candidates:
        if c and Path(c).exists():
            return c
    raise FileNotFoundError("LibreOffice not found; install it or set SOFFICE to soffice's path")
