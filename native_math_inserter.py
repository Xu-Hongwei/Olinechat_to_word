from __future__ import annotations

import argparse
import json
import re
import shutil
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": WORD_NS, "m": MATH_NS}
W = f"{{{WORD_NS}}}"
M = f"{{{MATH_NS}}}"

ET.register_namespace("w", WORD_NS)
ET.register_namespace("m", MATH_NS)

COMMAND_TEXT = {
    "alpha": "\u03b1",
    "beta": "\u03b2",
    "gamma": "\u03b3",
    "delta": "\u03b4",
    "epsilon": "\u03b5",
    "varepsilon": "\u03b5",
    "zeta": "\u03b6",
    "eta": "\u03b7",
    "theta": "\u03b8",
    "vartheta": "\u03d1",
    "iota": "\u03b9",
    "kappa": "\u03ba",
    "lambda": "\u03bb",
    "mu": "\u03bc",
    "nu": "\u03bd",
    "xi": "\u03be",
    "pi": "\u03c0",
    "rho": "\u03c1",
    "sigma": "\u03c3",
    "varsigma": "\u03c2",
    "tau": "\u03c4",
    "phi": "\u03c6",
    "varphi": "\u03d5",
    "chi": "\u03c7",
    "psi": "\u03c8",
    "omega": "\u03c9",
    "Gamma": "\u0393",
    "Delta": "\u0394",
    "Theta": "\u0398",
    "Lambda": "\u039b",
    "Xi": "\u039e",
    "Pi": "\u03a0",
    "Sigma": "\u03a3",
    "Phi": "\u03a6",
    "Psi": "\u03a8",
    "Omega": "\u03a9",
    "cdot": "\u00b7",
    "times": "\u00d7",
    "div": "\u00f7",
    "pm": "\u00b1",
    "mp": "\u2213",
    "neq": "\u2260",
    "ne": "\u2260",
    "leq": "\u2264",
    "le": "\u2264",
    "geq": "\u2265",
    "ge": "\u2265",
    "infty": "\u221e",
    "rightarrow": "\u2192",
    "leftarrow": "\u2190",
    "leftrightarrow": "\u2194",
    "Rightarrow": "\u21d2",
    "Leftarrow": "\u21d0",
    "Leftrightarrow": "\u21d4",
    "mapsto": "\u21a6",
    "to": "\u2192",
    "approx": "\u2248",
    "sim": "\u223c",
    "sum": "\u2211",
    "prod": "\u220f",
    "int": "\u222b",
    "partial": "\u2202",
    "nabla": "\u2207",
    "forall": "\u2200",
    "exists": "\u2203",
    "neg": "\u00ac",
    "in": "\u2208",
    "notin": "\u2209",
    "subseteq": "\u2286",
    "subset": "\u2282",
    "supseteq": "\u2287",
    "supset": "\u2283",
    "cup": "\u222a",
    "cap": "\u2229",
    "emptyset": "\u2205",
    "setminus": "\u2216",
    "parallel": "\u2225",
    "Vert": "\u2016",
    "vert": "|",
    "mid": "|",
    "ldots": "\u2026",
    "cdots": "\u22ef",
    "dots": "\u2026",
    "sin": "sin",
    "cos": "cos",
    "tan": "tan",
    "ln": "ln",
    "log": "log",
    "min": "min",
    "max": "max",
    "ell": "\u2113",
    "bigcup": "\u222a",
    "bigcap": "\u2229",
    "top": "\u22a4",
}
STYLE_COMMANDS = {"mathcal": ("scr", "script"), "mathbb": ("scr", "double-struck"), "boldsymbol": ("sty", "b")}
COMMAND_TEXT.update({
    "leqslant": "\u2264",
    "geqslant": "\u2265",
    "leqq": "\u2266",
    "geqq": "\u2267",
    "ll": "\u226a",
    "gg": "\u226b",
    "nexists": "\u2204",
    "ni": "\u220b",
    "owns": "\u220b",
    "simeq": "\u2243",
    "doteq": "\u2250",
    "triangleq": "\u225c",
    "gets": "\u2190",
    "longrightarrow": "\u2192",
    "longleftarrow": "\u2190",
    "longleftrightarrow": "\u2194",
    "Longrightarrow": "\u21d2",
    "Longleftarrow": "\u21d0",
    "Longleftrightarrow": "\u21d4",
    "iff": "\u21d4",
    "implies": "\u21d2",
    "impliedby": "\u21d0",
    "sup": "sup",
    "inf": "inf",
    "argmax": "argmax",
    "argmin": "argmin",
    "equiv": "\u2261",
    "propto": "\u221d",
    "cong": "\u2245",
    "oint": "\u222e",
    "land": "\u2227",
    "wedge": "\u2227",
    "lor": "\u2228",
    "vee": "\u2228",
    "subsetneq": "\u228a",
    "supsetneq": "\u228b",
    "bigvee": "\u22c1",
    "bigwedge": "\u22c0",
    "perp": "\u22a5",
    "therefore": "\u2234",
    "because": "\u2235",
    "angle": "\u2220",
    "triangle": "\u25b3",
    "degree": "\u00b0",
    "cot": "cot",
    "sec": "sec",
    "csc": "csc",
    "lim": "lim",
    "limsup": "lim sup",
    "liminf": "lim inf",
    "arg": "arg",
    "det": "det",
    "gcd": "gcd",
    "Pr": "Pr",
    "bot": "\u22a5",
    "Re": "\u211c",
    "Im": "\u2111",
})

OPEN_DELIMS = {"(": ")", "[": "]", r"\{": r"\}", "|": "|", "\u2016": "\u2016"}
DELIM_RENDER = {r"\{": "{", r"\}": "}", "(": "(", ")": ")", "[": "[", "]": "]", "|": "|", "\u2016": "\u2016"}

ACCENT_COMMANDS = {
    "hat": "\u0302",
    "widehat": "\u0302",
    "bar": "\u0304",
    "overline": "\u0305",
    "tilde": "\u0303",
    "widetilde": "\u0303",
    "vec": "\u20d7",
    "overrightarrow": "\u20d7",
    "dot": "\u0307",
    "ddot": "\u0308",
    "check": "\u030c",
    "breve": "\u0306",
}
LIMIT_STYLE_BASES = {
    "max",
    "min",
    "lim",
    "lim sup",
    "lim inf",
    "sup",
    "inf",
    "argmax",
    "argmin",
    "\u2211",
    "\u220f",
    "\u222b",
    "\u222e",
    "\u222a",
    "\u2229",
    "\u22c1",
    "\u22c0",
}


@dataclass
class Token:
    kind: str
    value: str


@dataclass
class SequenceNode:
    items: list


@dataclass
class SymbolNode:
    text: str
    style: tuple[str, str] | None = None


@dataclass
class DelimitedNode:
    left: str
    right: str
    content: SequenceNode


@dataclass
class FractionNode:
    numerator: SequenceNode
    denominator: SequenceNode


@dataclass
class RadicalNode:
    radicand: SequenceNode
    degree: SequenceNode | None = None


@dataclass
class MatrixNode:
    rows: list[list[SequenceNode]]
    left: str | None = None
    right: str | None = None


@dataclass
class AccentNode:
    accent: str
    base: SequenceNode


@dataclass
class LimitNode:
    base: object
    lower: SequenceNode | None = None
    upper: SequenceNode | None = None


@dataclass
class ScriptNode:
    base: object
    sub: SequenceNode | None = None
    sup: SequenceNode | None = None


class LatexTokenizer:
    def __init__(self, text: str) -> None:
        self.text = text
        self.length = len(text)
        self.index = 0

    def tokenize(self) -> list[Token]:
        tokens: list[Token] = []
        while self.index < self.length:
            char = self.text[self.index]
            if char.isspace():
                self.index += 1
                continue
            if char == "\\":
                if self.index + 1 < self.length and self.text[self.index + 1] == "\\":
                    tokens.append(Token("row_sep", "\\\\"))
                    self.index += 2
                    continue
                if self.index + 1 < self.length and self.text[self.index + 1] in "{}[]()|":
                    tokens.append(Token("delim", self.text[self.index:self.index + 2]))
                    self.index += 2
                    continue
                match = re.match(r"\\[A-Za-z]+", self.text[self.index:])
                if match:
                    tokens.append(Token("command", match.group(0)[1:]))
                    self.index += len(match.group(0))
                    continue
                tokens.append(Token("text", char))
                self.index += 1
                continue
            if char in "{}_^()[],:=+-*|&":
                tokens.append(Token("char", char))
                self.index += 1
                continue
            if char == "?":
                tokens.append(Token("char", char))
                self.index += 1
                continue
            if char.isdigit():
                start = self.index
                while self.index < self.length and self.text[self.index].isdigit():
                    self.index += 1
                tokens.append(Token("text", self.text[start:self.index]))
                continue
            if char.isalpha():
                tokens.append(Token("text", char))
                self.index += 1
                continue
            tokens.append(Token("text", char))
            self.index += 1
        return tokens


class LatexParser:
    def __init__(self, text: str) -> None:
        self.tokens = LatexTokenizer(text).tokenize()
        self.index = 0

    def parse(self) -> SequenceNode:
        return self.parse_sequence(stop_values=set())

    def peek(self) -> Token | None:
        if self.index >= len(self.tokens):
            return None
        return self.tokens[self.index]

    def consume(self) -> Token:
        token = self.tokens[self.index]
        self.index += 1
        return token

    def parse_sequence(self, stop_values: set[str]) -> SequenceNode:
        items: list[object] = []
        while self.index < len(self.tokens):
            token = self.peek()
            if token is None:
                break
            if token.value in stop_values:
                break
            items.append(self.parse_atom())
        return SequenceNode(items)

    def parse_required_group(self) -> SequenceNode:
        token = self.peek()
        if token is None:
            return SequenceNode([])
        if token.kind == "char" and token.value == "{":
            self.consume()
            content = self.parse_sequence({"}"})
            if self.peek() and self.peek().value == "}":
                self.consume()
            return content
        return SequenceNode([self.parse_atom()])

    def parse_script_arg(self) -> SequenceNode:
        token = self.peek()
        if token is None:
            return SequenceNode([])
        if token.kind == "char" and token.value == "{":
            return self.parse_required_group()
        return SequenceNode([self.parse_primary()])

    def parse_optional_group(self, open_char: str, close_char: str) -> SequenceNode | None:
        token = self.peek()
        if token is None or token.kind != "char" or token.value != open_char:
            return None
        self.consume()
        content = self.parse_sequence({close_char})
        if self.peek() and self.peek().value == close_char:
            self.consume()
        return content

    def sequence_to_text(self, sequence: SequenceNode) -> str:
        parts: list[str] = []
        for item in sequence.items:
            if isinstance(item, SymbolNode):
                parts.append(item.text)
            elif isinstance(item, SequenceNode):
                parts.append(self.sequence_to_text(item))
            else:
                parts.append(str(item))
        return "".join(parts)

    def matches_environment_end(self, env_name: str) -> bool:
        if self.index >= len(self.tokens):
            return False
        token = self.tokens[self.index]
        if token.kind != "command" or token.value != "end":
            return False
        cursor = self.index + 1
        if cursor >= len(self.tokens) or self.tokens[cursor].value != "{":
            return False
        cursor += 1
        parts: list[str] = []
        while cursor < len(self.tokens) and self.tokens[cursor].value != "}":
            parts.append(self.tokens[cursor].value)
            cursor += 1
        if cursor >= len(self.tokens):
            return False
        return "".join(parts) == env_name

    def consume_environment_end(self, env_name: str) -> None:
        self.consume()
        end_name = self.sequence_to_text(self.parse_required_group())
        if end_name != env_name:
            raise ValueError(f"Expected \\end{{{env_name}}}, got \\end{{{end_name}}}")

    def is_empty_sequence(self, sequence: SequenceNode) -> bool:
        if not sequence.items:
            return True
        for item in sequence.items:
            if isinstance(item, SymbolNode) and item.text == "":
                continue
            if isinstance(item, SequenceNode) and self.is_empty_sequence(item):
                continue
            return False
        return True

    def parse_environment_body(self, env_name: str) -> list[list[SequenceNode]]:
        if env_name == "array" and self.peek() and self.peek().value == "{":
            self.parse_required_group()

        rows: list[list[SequenceNode]] = []
        current_row: list[SequenceNode] = []
        current_cell_items: list[object] = []

        while self.index < len(self.tokens):
            if self.matches_environment_end(env_name):
                self.consume_environment_end(env_name)
                break

            token = self.peek()
            if token is None:
                break
            if token.kind == "char" and token.value == "&":
                self.consume()
                current_row.append(SequenceNode(current_cell_items))
                current_cell_items = []
                continue
            if token.kind == "row_sep":
                self.consume()
                current_row.append(SequenceNode(current_cell_items))
                current_cell_items = []
                rows.append(current_row)
                current_row = []
                continue
            current_cell_items.append(self.parse_atom())

        current_row.append(SequenceNode(current_cell_items))
        rows.append(current_row)

        while rows and len(rows[-1]) == 1 and self.is_empty_sequence(rows[-1][0]):
            rows.pop()
        return rows or [[SequenceNode([])]]

    def parse_environment(self, env_name: str):
        delimiter_environments = {
            "pmatrix": ("(", ")"),
            "bmatrix": ("[", "]"),
            "Bmatrix": ("{", "}"),
            "vmatrix": ("|", "|"),
            "Vmatrix": (chr(0x2016), chr(0x2016)),
            "cases": ("{", ""),
        }
        if env_name in delimiter_environments:
            left, right = delimiter_environments[env_name]
            return MatrixNode(self.parse_environment_body(env_name), left=left, right=right)
        if env_name in {"matrix", "smallmatrix", "array", "aligned", "align", "alignedat", "gathered", "split"}:
            return MatrixNode(self.parse_environment_body(env_name))
        return SymbolNode(f"\\begin{{{env_name}}}")

    def parse_primary(self):
        token = self.peek()
        if token is None:
            return SymbolNode("")

        if token.kind == "char" and token.value == "{":
            return self.parse_required_group()

        if token.kind in {"char", "delim"} and token.value in OPEN_DELIMS:
            open_token = self.consume().value
            close_token = OPEN_DELIMS[open_token]
            content = self.parse_sequence({close_token})
            if self.peek() and self.peek().value == close_token:
                self.consume()
            return DelimitedNode(DELIM_RENDER[open_token], DELIM_RENDER[close_token], content)

        if token.kind == "command":
            command = self.consume().value
            if command == "begin":
                env_name = self.sequence_to_text(self.parse_required_group())
                return self.parse_environment(env_name)
            if command in STYLE_COMMANDS:
                content = self.parse_required_group()
                if len(content.items) == 1 and isinstance(content.items[0], SymbolNode):
                    return SymbolNode(content.items[0].text, style=STYLE_COMMANDS[command])
                content_items = []
                for item in content.items:
                    if isinstance(item, SymbolNode):
                        content_items.append(SymbolNode(item.text, style=STYLE_COMMANDS[command]))
                    else:
                        content_items.append(item)
                return SequenceNode(content_items)
            if command in {"frac", "dfrac", "tfrac", "cfrac"}:
                return FractionNode(self.parse_required_group(), self.parse_required_group())
            if command in {"binom", "dbinom", "tbinom"}:
                return MatrixNode(
                    [[self.parse_required_group()], [self.parse_required_group()]],
                    left="(",
                    right=")",
                )
            if command == "sqrt":
                degree = self.parse_optional_group("[", "]")
                return RadicalNode(
                    self.parse_required_group(),
                    degree=degree if degree and degree.items else None,
                )
            if command in ACCENT_COMMANDS:
                return AccentNode(ACCENT_COMMANDS[command], self.parse_required_group())
            if command in {"xrightarrow", "xleftarrow", "xleftrightarrow", "xRightarrow", "xLeftarrow", "xLeftrightarrow"}:
                lower = self.parse_optional_group("[", "]")
                upper = self.parse_required_group()
                arrow_text = {
                    "xrightarrow": "\u2192",
                    "xleftarrow": "\u2190",
                    "xleftrightarrow": "\u2194",
                    "xRightarrow": "\u21d2",
                    "xLeftarrow": "\u21d0",
                    "xLeftrightarrow": "\u21d4",
                }[command]
                arrow = SymbolNode(arrow_text)
                return LimitNode(arrow, lower=lower if lower and lower.items else None, upper=upper if upper and upper.items else None)
            return SymbolNode(COMMAND_TEXT.get(command, "\\" + command))

        return SymbolNode(self.consume().value)

    def parse_atom(self):
        return self.apply_scripts(self.parse_primary())

    def apply_scripts(self, node):
        sub = None
        sup = None
        while self.peek() and self.peek().value in {"_", "^"}:
            marker = self.consume().value
            arg = self.parse_script_arg()
            if marker == "_":
                sub = arg
            else:
                sup = arg
        if sub is not None or sup is not None:
            return ScriptNode(node, sub=sub, sup=sup)
        return node

class OmmlBuilder:
    def create_run(self, text: str, style: tuple[str, str] | None = None) -> ET.Element:
        run = ET.Element(f"{M}r")
        if style:
            style_tag, style_value = style
            rpr = ET.SubElement(run, f"{M}rPr")
            style_element = ET.SubElement(rpr, f"{M}{style_tag}")
            style_element.set(f"{M}val", style_value)
        text_element = ET.SubElement(run, f"{M}t")
        text_element.text = text
        return run

    def append_sequence(self, parent: ET.Element, sequence: SequenceNode) -> None:
        for item in sequence.items:
            for element in self.build(item):
                parent.append(element)

    def make_container(self, tag: str, sequence: SequenceNode) -> ET.Element:
        container = ET.Element(tag)
        self.append_sequence(container, sequence)
        return container

    def build_limit(self, base_node, *, lower: SequenceNode | None = None, upper: SequenceNode | None = None) -> list[ET.Element]:
        base_elements = self.build(base_node)
        if len(base_elements) == 1:
            current = base_elements[0]
        else:
            wrapper = ET.Element(f"{M}box")
            e = ET.SubElement(wrapper, f"{M}e")
            for element in base_elements:
                e.append(element)
            current = wrapper
        if lower is not None:
            lim_low = ET.Element(f"{M}limLow")
            e = ET.SubElement(lim_low, f"{M}e")
            e.append(current)
            lim = ET.SubElement(lim_low, f"{M}lim")
            self.append_sequence(lim, lower)
            current = lim_low
        if upper is not None:
            lim_upp = ET.Element(f"{M}limUpp")
            e = ET.SubElement(lim_upp, f"{M}e")
            e.append(current)
            lim = ET.SubElement(lim_upp, f"{M}lim")
            self.append_sequence(lim, upper)
            current = lim_upp
        return [current]

    def build(self, node) -> list[ET.Element]:
        if isinstance(node, SequenceNode):
            result: list[ET.Element] = []
            for item in node.items:
                result.extend(self.build(item))
            return result
        if isinstance(node, SymbolNode):
            return [self.create_run(node.text, style=node.style)]
        if isinstance(node, MatrixNode):
            matrix = ET.Element(f"{M}m")
            for row in node.rows:
                matrix_row = ET.SubElement(matrix, f"{M}mr")
                for cell in row:
                    cell_element = ET.SubElement(matrix_row, f"{M}e")
                    self.append_sequence(cell_element, cell)
            if node.left is None and node.right is None:
                return [matrix]
            delimiter = ET.Element(f"{M}d")
            dpr = ET.SubElement(delimiter, f"{M}dPr")
            beg = ET.SubElement(dpr, f"{M}begChr")
            beg.set(f"{M}val", node.left or "")
            end = ET.SubElement(dpr, f"{M}endChr")
            end.set(f"{M}val", node.right or "")
            content = ET.SubElement(delimiter, f"{M}e")
            content.append(matrix)
            return [delimiter]
        if isinstance(node, DelimitedNode):
            delimiter = ET.Element(f"{M}d")
            dpr = ET.SubElement(delimiter, f"{M}dPr")
            beg = ET.SubElement(dpr, f"{M}begChr")
            beg.set(f"{M}val", node.left)
            end = ET.SubElement(dpr, f"{M}endChr")
            end.set(f"{M}val", node.right)
            content = ET.SubElement(delimiter, f"{M}e")
            self.append_sequence(content, node.content)
            return [delimiter]
        if isinstance(node, FractionNode):
            fraction = ET.Element(f"{M}f")
            numerator = ET.SubElement(fraction, f"{M}num")
            self.append_sequence(numerator, node.numerator)
            denominator = ET.SubElement(fraction, f"{M}den")
            self.append_sequence(denominator, node.denominator)
            return [fraction]
        if isinstance(node, RadicalNode):
            radical = ET.Element(f"{M}rad")
            radical_pr = ET.SubElement(radical, f"{M}radPr")
            if node.degree is None:
                deg_hide = ET.SubElement(radical_pr, f"{M}degHide")
                deg_hide.set(f"{M}val", "1")
            else:
                degree = ET.SubElement(radical, f"{M}deg")
                self.append_sequence(degree, node.degree)
            base = ET.SubElement(radical, f"{M}e")
            self.append_sequence(base, node.radicand)
            return [radical]
        if isinstance(node, AccentNode):
            accent = ET.Element(f"{M}acc")
            accent_pr = ET.SubElement(accent, f"{M}accPr")
            chr_element = ET.SubElement(accent_pr, f"{M}chr")
            chr_element.set(f"{M}val", node.accent)
            base = ET.SubElement(accent, f"{M}e")
            self.append_sequence(base, node.base)
            return [accent]
        if isinstance(node, LimitNode):
            return self.build_limit(node.base, lower=node.lower, upper=node.upper)
        if isinstance(node, ScriptNode):
            if isinstance(node.base, SymbolNode) and node.base.text in LIMIT_STYLE_BASES:
                return self.build_limit(node.base, lower=node.sub, upper=node.sup)
            tag = f"{M}sSubSup" if node.sub is not None and node.sup is not None else (f"{M}sSub" if node.sub is not None else f"{M}sSup")
            script = ET.Element(tag)
            base = ET.SubElement(script, f"{M}e")
            for element in self.build(node.base):
                base.append(element)
            if node.sub is not None:
                sub = ET.SubElement(script, f"{M}sub")
                self.append_sequence(sub, node.sub)
            if node.sup is not None:
                sup = ET.SubElement(script, f"{M}sup")
                self.append_sequence(sup, node.sup)
            return [script]
        return [self.create_run(str(node))]

    def build_omath(self, expression: str) -> ET.Element:
        parser = LatexParser(expression)
        sequence = parser.parse()
        omath = ET.Element(f"{M}oMath")
        self.append_sequence(omath, sequence)
        return omath


def build_parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    parent_map: dict[ET.Element, ET.Element] = {}
    for parent in root.iter():
        for child in list(parent):
            parent_map[child] = parent
    return parent_map


def build_omml_fragment(expression: str, *, display: bool = False) -> str:
    builder = OmmlBuilder()
    omath = builder.build_omath(expression)
    if display:
        omath_para = ET.Element(f"{M}oMathPara")
        omath_para.append(omath)
        fragment = ET.tostring(omath_para, encoding="unicode")
    else:
        fragment = ET.tostring(omath, encoding="unicode")
    fragment = fragment.replace(f' xmlns:m="{MATH_NS}"', "")
    return fragment


def build_placeholder_run_pattern(placeholder: str) -> re.Pattern[str]:
    escaped = re.escape(placeholder)
    return re.compile(
        rf"<w:r\b[^>]*>(?:(?!<w:r\b).)*?"
        rf"<w:t(?:\s+xml:space=\"preserve\")?>{escaped}</w:t>"
        rf"(?:(?!<w:r\b).)*?</w:r>",
        re.DOTALL,
    )


def replace_math_placeholders(docx_path: Path, payload_path: Path) -> int:
    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    math_items = payload.get("math_items", [])

    with ZipFile(docx_path, "r") as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")

    replacements = 0
    for item in math_items:
        placeholder = item["placeholder"]
        expression = item["text"]
        display = bool(item.get("display", False))
        fragment = build_omml_fragment(expression, display=display)
        pattern = build_placeholder_run_pattern(placeholder)
        document_xml, count = pattern.subn(lambda _match: fragment, document_xml, count=1)
        replacements += count

    temp_path = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with ZipFile(docx_path, "r") as source, ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = document_xml.encode("utf-8") if info.filename == "word/document.xml" else source.read(info.filename)
            target.writestr(info, data)
    shutil.move(temp_path, docx_path)
    return replacements

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-json", required=True)
    parser.add_argument("--input-docx", required=True)
    args = parser.parse_args()

    count = replace_math_placeholders(Path(args.input_docx), Path(args.input_json))
    print(f"Native OMML inserted: {count}")


if __name__ == "__main__":
    main()
