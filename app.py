from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx_layout_builder import build_document
from native_math_inserter import replace_math_placeholders


COMMAND_MAP = {
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ε",
    r"\varepsilon": "ε",
    r"\zeta": "ζ",
    r"\eta": "η",
    r"\theta": "θ",
    r"\vartheta": "ϑ",
    r"\iota": "ι",
    r"\kappa": "κ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\nu": "ν",
    r"\xi": "ξ",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\varsigma": "ς",
    r"\tau": "τ",
    r"\phi": "φ",
    r"\varphi": "ϕ",
    r"\chi": "χ",
    r"\psi": "ψ",
    r"\omega": "ω",
    r"\Gamma": "Γ",
    r"\Delta": "Δ",
    r"\Theta": "Θ",
    r"\Lambda": "Λ",
    r"\Xi": "Ξ",
    r"\Pi": "Π",
    r"\Sigma": "Σ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
    r"\cdot": "·",
    r"\times": "×",
    r"\div": "÷",
    r"\pm": "±",
    r"\mp": "∓",
    r"\neq": "≠",
    r"\ne": "≠",
    r"\leq": "≤",
    r"\le": "≤",
    r"\geq": "≥",
    r"\ge": "≥",
    r"\infty": "∞",
    r"\rightarrow": "→",
    r"\leftarrow": "←",
    r"\leftrightarrow": "↔",
    r"\Rightarrow": "⇒",
    r"\Leftarrow": "⇐",
    r"\Leftrightarrow": "⇔",
    r"\mapsto": "↦",
    r"\to": "→",
    r"\approx": "≈",
    r"\sim": "∼",
    r"\sum": "∑",
    r"\prod": "∏",
    r"\int": "∫",
    r"\partial": "∂",
    r"\nabla": "∇",
    r"\forall": "∀",
    r"\exists": "∃",
    r"\neg": "¬",
    r"\in": "∈",
    r"\notin": "∉",
    r"\subseteq": "⊆",
    r"\subset": "⊂",
    r"\supseteq": "⊇",
    r"\supset": "⊃",
    r"\cup": "∪",
    r"\cap": "∩",
    r"\emptyset": "∅",
    r"\setminus": "∖",
    r"\parallel": "∥",
    r"\Vert": "‖",
    r"\vert": "|",
    r"\mid": "|",
    r"\ldots": "…",
    r"\cdots": "⋯",
    r"\dots": "…",
    r"\sin": "sin",
    r"\cos": "cos",
    r"\tan": "tan",
    r"\ln": "ln",
    r"\log": "log",
    r"\ell": "ℓ",
    r"\bigcup": "⋃",
    r"\bigcap": "⋂",
}

MATHCAL_CHAR_MAP = {
    "A": "𝒜", "B": "ℬ", "C": "𝒞", "D": "𝒟", "E": "ℰ", "F": "ℱ", "G": "𝒢", "H": "ℋ",
    "I": "ℐ", "J": "𝒥", "K": "𝒦", "L": "ℒ", "M": "ℳ", "N": "𝒩", "O": "𝒪", "P": "𝒫",
    "Q": "𝒬", "R": "ℛ", "S": "𝒮", "T": "𝒯", "U": "𝒰", "V": "𝒱", "W": "𝒲", "X": "𝒳",
    "Y": "𝒴", "Z": "𝒵",
}
MATHBB_CHAR_MAP = {
    "C": "ℂ", "H": "ℍ", "N": "ℕ", "P": "ℙ", "Q": "ℚ", "R": "ℝ", "Z": "ℤ",
    "A": "𝔸", "B": "𝔹", "D": "𝔻", "E": "𝔼", "F": "𝔽", "G": "𝔾", "I": "𝕀", "J": "𝕁",
    "K": "𝕂", "L": "𝕃", "M": "𝕄", "O": "𝕆", "S": "𝕊", "T": "𝕋", "U": "𝕌", "V": "𝕍",
    "W": "𝕎", "X": "𝕏", "Y": "𝕐",
}

INLINE_PATTERN = re.compile(r"(`[^`]+`|\\\([^\n]+?\\\)|\$\$.*?\$\$|\$[^$\n]+\$|\*\*[^*\n]+\*\*|\*[^*\n]+\*)")
TEXT_WRAPPER_COMMANDS = ("text", "mathrm", "mathbf", "operatorname", "mathit")
OPERATORNAME_COMMANDS = {
    "lim": r"\lim",
    "limsup": r"\limsup",
    "liminf": r"\liminf",
    "sup": r"\sup",
    "inf": r"\inf",
    "argmax": r"\argmax",
    "argmin": r"\argmin",
    "det": r"\det",
    "gcd": r"\gcd",
    "pr": r"\Pr",
}

BULLET_ITEM_PATTERN = re.compile(r"^(?P<indent>\s*)(?P<marker>[-*])\s+(?P<text>.+)$")
ORDERED_ITEM_PATTERN = re.compile(r"^(?P<indent>\s*)(?P<number>\d+)[.)]\s+(?P<text>.+)$")


def get_indent_level(indent: str) -> int:
    normalized = indent.replace("\t", "    ")
    return min(len(normalized) // 2, 8)


def style_math_text(content: str, mapping: dict[str, str]) -> str:
    normalized = normalize_equation_for_word(content)
    return "".join(mapping.get(character, character) for character in normalized)


def parse_list_item(line: str) -> dict | None:
    bullet_match = BULLET_ITEM_PATTERN.match(line)
    if bullet_match:
        return {
            "ordered": False,
            "level": get_indent_level(bullet_match.group("indent")),
            "runs": parse_inline(bullet_match.group("text").strip()),
        }
    ordered_match = ORDERED_ITEM_PATTERN.match(line)
    if ordered_match:
        return {
            "ordered": True,
            "level": get_indent_level(ordered_match.group("indent")),
            "start": int(ordered_match.group("number")),
            "runs": parse_inline(ordered_match.group("text").strip()),
        }
    return None


def normalize_markdown_content(markdown_text: str) -> str:
    normalized = markdown_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    result: list[str] = []
    in_code_block = False
    blank_pending = False

    for line in lines:
        current = line.rstrip()
        stripped = current.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            result.append(current)
            blank_pending = False
            continue

        if in_code_block:
            result.append(current)
            blank_pending = False
            continue

        if not stripped:
            if not blank_pending:
                result.append("")
            blank_pending = True
            continue

        if re.match(r"^[-*_]{2,}$", stripped):
            current = "---"
        else:
            heading_match = re.match(r"^(#{1,6})\s*(.+)$", stripped)
            if heading_match:
                current = f"{heading_match.group(1)} {heading_match.group(2).strip()}"
                if result and result[-1] != "":
                    result.append("")
            else:
                bullet_match = re.match(r"^(\s*)([-*])\s+(.+)$", current)
                if bullet_match and bullet_match.group(3).strip():
                    current = f"{bullet_match.group(1)}{bullet_match.group(2)} {bullet_match.group(3).strip()}"
                else:
                    ordered_match = re.match(r"^(\s*)(\d+)[.)]\s*(.+)$", current)
                    if ordered_match and ordered_match.group(3).strip():
                        current = f"{ordered_match.group(1)}{ordered_match.group(2)}. {ordered_match.group(3).strip()}"
                    else:
                        current = stripped

        result.append(current)
        blank_pending = False

    while result and result[-1] == "":
        result.pop()
    return "\n".join(result)


def split_braced(text: str, start_index: int) -> tuple[str, int] | None:
    if start_index >= len(text) or text[start_index] != "{":
        return None
    depth = 0
    for index in range(start_index, len(text)):
        character = text[index]
        if character == "{":
            depth += 1
        elif character == "}":
            depth -= 1
            if depth == 0:
                return text[start_index + 1:index], index + 1
    return None


def replace_latex_func(text: str, command: str, repl) -> str:
    needle = f"\\{command}"
    cursor = 0
    parts: list[str] = []
    while cursor < len(text):
        position = text.find(needle, cursor)
        if position == -1:
            parts.append(text[cursor:])
            break
        parts.append(text[cursor:position])
        arg = split_braced(text, position + len(needle))
        if not arg:
            parts.append(needle)
            cursor = position + len(needle)
            continue
        content, next_index = arg
        parts.append(repl(content))
        cursor = next_index
    return "".join(parts)


def replace_frac(text: str) -> str:
    cursor = 0
    parts: list[str] = []
    while cursor < len(text):
        position = text.find(r"\frac", cursor)
        if position == -1:
            parts.append(text[cursor:])
            break
        parts.append(text[cursor:position])
        first = split_braced(text, position + len(r"\frac"))
        if not first:
            parts.append(r"\frac")
            cursor = position + len(r"\frac")
            continue
        numerator, next_index = first
        second = split_braced(text, next_index)
        if not second:
            parts.append(r"\frac{" + numerator + "}")
            cursor = next_index
            continue
        denominator, cursor = second
        parts.append(f"({normalize_equation_for_word(numerator)})/({normalize_equation_for_word(denominator)})")
    return "".join(parts)


def replace_extended_arrows(text: str) -> str:
    return text


def replace_spacing_commands(text: str) -> str:
    replacements = {
        r"\qquad": " ",
        r"\quad": " ",
        r"\;": " ",
        r"\,": " ",
        r"\!": "",
    }
    value = text
    for source, target in replacements.items():
        value = value.replace(source, target)
    return value


def has_outer_group(text: str) -> bool:
    if len(text) < 2 or text[0] != "(" or text[-1] != ")":
        return False
    depth = 0
    for index, character in enumerate(text):
        if character == "(":
            depth += 1
        elif character == ")":
            depth -= 1
            if depth == 0 and index != len(text) - 1:
                return False
    return depth == 0


def format_script_value(marker: str, content: str) -> str:
    normalized = normalize_equation_for_word(content).strip()
    if not normalized:
        return marker
    if has_outer_group(normalized):
        return f"{marker}{normalized}"
    if all(character.isalnum() or character in {"*", "?"} for character in normalized):
        return f"{marker}{normalized}"
    return f"{marker}({normalized})"


def replace_script_braces(text: str) -> str:
    cursor = 0
    parts: list[str] = []
    while cursor < len(text):
        marker = text[cursor]
        if marker not in {"_", "^"}:
            parts.append(marker)
            cursor += 1
            continue
        cursor += 1
        if cursor >= len(text):
            parts.append(marker)
            break
        if text[cursor] == "(":
            parts.append(marker)
            continue
        if text[cursor] == "{":
            arg = split_braced(text, cursor)
            if arg:
                content, cursor = arg
                parts.append(format_script_value(marker, content))
                continue
        if text[cursor] == "\\":
            command_match = re.match(r"\\[A-Za-z]+", text[cursor:])
            if command_match:
                token = command_match.group(0)
                cursor += len(token)
                parts.append(format_script_value(marker, token))
                continue
        parts.append(format_script_value(marker, text[cursor]))
        cursor += 1
    return "".join(parts)


def normalize_unbraced_style_command(text: str, command: str) -> str:
    backslash = re.escape("\\")
    pattern = re.compile(rf"{backslash}{command}(?P<token>{backslash}[A-Za-z]+|[A-Za-z0-9])(?=[^A-Za-z0-9]|$)")

    def repl(match: re.Match[str]) -> str:
        token = match.group("token")
        normalized = normalize_equation_for_word(token)
        return "\\" + command + "{" + normalized + "}"

    return pattern.sub(repl, text)

def normalize_preserved_command(text: str, command: str) -> str:
    needle = f"\\{command}"
    cursor = 0
    parts: list[str] = []
    while cursor < len(text):
        position = text.find(needle, cursor)
        if position == -1:
            parts.append(text[cursor:])
            break
        parts.append(text[cursor:position])
        arg = split_braced(text, position + len(needle))
        if not arg:
            parts.append(needle)
            cursor = position + len(needle)
            continue
        content, next_index = arg
        normalized_content = normalize_equation_for_word(content)
        parts.append(f"\\{command}" + "{" + normalized_content + "}")
        cursor = next_index
    return "".join(parts)


def normalize_operator_name(content: str) -> str:
    normalized = normalize_equation_for_word(content).strip()
    collapsed = re.sub(r"\s+", "", normalized).lower()
    return OPERATORNAME_COMMANDS.get(collapsed, normalized)


def normalize_equation_for_word(text: str) -> str:
    value = text.strip().replace("\n", " ")
    value = replace_spacing_commands(value)
    value = replace_extended_arrows(value)
    value = value.replace(r"\operatorname*", r"\operatorname")
    for source, target in (
        (r"\dfrac", r"\frac"),
        (r"\tfrac", r"\frac"),
        (r"\cfrac", r"\frac"),
    ):
        value = value.replace(source, target)
    for source, target in (
        (r"\gets", r"\leftarrow"),
        (r"\longrightarrow", r"\rightarrow"),
        (r"\longleftarrow", r"\leftarrow"),
        (r"\longleftrightarrow", r"\leftrightarrow"),
        (r"\Longrightarrow", r"\Rightarrow"),
        (r"\Longleftarrow", r"\Leftarrow"),
        (r"\Longleftrightarrow", r"\Leftrightarrow"),
        (r"\iff", r"\Leftrightarrow"),
        (r"\implies", r"\Rightarrow"),
        (r"\impliedby", r"\Leftarrow"),
        (r"\leqslant", r"\leq"),
        (r"\geqslant", r"\geq"),
        (r"\leqq", r"\leq"),
        (r"\geqq", r"\geq"),
    ):
        value = value.replace(source, target)
    for source, target in (
        (r"\left(", "("),
        (r"\right)", ")"),
        (r"\left[", "["),
        (r"\right]", "]"),
        (r"\left\{", r"\{"),
        (r"\right\}", r"\}"),
        (r"\left|", "|"),
        (r"\right|", "|"),
        (r"\left\lvert", "|"),
        (r"\right\rvert", "|"),
        (r"\left\Vert", r"\Vert"),
        (r"\right\Vert", r"\Vert"),
    ):
        value = value.replace(source, target)
    value = normalize_unbraced_style_command(value, "boldsymbol")
    value = normalize_unbraced_style_command(value, "hat")
    value = normalize_preserved_command(value, "mathcal")
    value = normalize_preserved_command(value, "mathbb")
    value = normalize_preserved_command(value, "boldsymbol")
    value = normalize_preserved_command(value, "hat")
    for command in ("mathbf", "mathrm", "mathit", "text"):
        value = replace_latex_func(value, command, lambda arg: normalize_equation_for_word(arg))
    value = replace_latex_func(value, "operatorname", normalize_operator_name)
    value = value.replace(r"\bigcup", r"\cup")
    value = value.replace(r"\bigcap", r"\cap")
    value = re.sub(r"\s+", " ", value)
    return value.strip()

def parse_inline(text: str) -> list[dict[str, str]]:
    runs: list[dict[str, str]] = []
    cursor = 0
    for match in INLINE_PATTERN.finditer(text):
        start, end = match.span()
        if start > cursor:
            runs.append({"type": "text", "text": text[cursor:start]})
        token = match.group(0)
        if token.startswith("\\(") and token.endswith("\\)"):
            runs.append({"type": "math", "text": normalize_equation_for_word(token[2:-2])})
        elif token.startswith("$$") and token.endswith("$$"):
            runs.append({"type": "math", "text": normalize_equation_for_word(token[2:-2])})
        elif token.startswith("$") and token.endswith("$"):
            runs.append({"type": "math", "text": normalize_equation_for_word(token[1:-1])})
        elif token.startswith("**") and token.endswith("**"):
            runs.append({"type": "bold", "text": token[2:-2]})
        elif token.startswith("*") and token.endswith("*"):
            runs.append({"type": "italic", "text": token[1:-1]})
        elif token.startswith("`") and token.endswith("`"):
            runs.append({"type": "code", "text": token[1:-1]})
        cursor = end
    if cursor < len(text):
        runs.append({"type": "text", "text": text[cursor:]})
    return [run for run in runs if run["text"]]


def parse_markdown(markdown_text: str) -> list[dict]:
    normalized = markdown_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = normalized.split("\n")
    blocks: list[dict] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue
        if re.match(r"^\s*---+\s*$", stripped):
            blocks.append({"type": "separator"})
            index += 1
            continue
        if stripped == r"\[":
            buffer = []
            index += 1
            while index < len(lines) and lines[index].strip() != r"\]":
                buffer.append(lines[index])
                index += 1
            if index < len(lines):
                index += 1
            blocks.append({"type": "math_block", "text": normalize_equation_for_word("\n".join(buffer))})
            continue
        if stripped.startswith(r"\[") and stripped.endswith(r"\]") and len(stripped) > 4:
            blocks.append({"type": "math_block", "text": normalize_equation_for_word(stripped[2:-2])})
            index += 1
            continue
        if stripped.startswith("```"):
            language = stripped[3:].strip()
            buffer: list[str] = []
            index += 1
            while index < len(lines) and not lines[index].strip().startswith("```"):
                buffer.append(lines[index])
                index += 1
            if index < len(lines):
                index += 1
            blocks.append({"type": "code_block", "language": language, "text": "\n".join(buffer)})
            continue
        if stripped == "$$":
            buffer = []
            index += 1
            while index < len(lines) and lines[index].strip() != "$$":
                buffer.append(lines[index])
                index += 1
            if index < len(lines):
                index += 1
            blocks.append({"type": "math_block", "text": normalize_equation_for_word("\n".join(buffer))})
            continue
        if stripped.startswith("$$") and stripped.endswith("$$") and len(stripped) > 4:
            blocks.append({"type": "math_block", "text": normalize_equation_for_word(stripped[2:-2])})
            index += 1
            continue
        heading_match = re.match(r"^(#{1,6})\s+(.*)$", line)
        if heading_match:
            blocks.append({
                "type": "heading",
                "level": len(heading_match.group(1)),
                "runs": parse_inline(heading_match.group(2).strip()),
            })
            index += 1
            continue
        quote_match = re.match(r"^\s*>\s?(.*)$", line)
        if quote_match:
            buffer = [quote_match.group(1)]
            index += 1
            while index < len(lines):
                next_match = re.match(r"^\s*>\s?(.*)$", lines[index])
                if not next_match:
                    break
                buffer.append(next_match.group(1))
                index += 1
            blocks.append({"type": "blockquote", "runs": parse_inline(" ".join(part.strip() for part in buffer if part.strip()))})
            continue
        list_item = parse_list_item(line)
        if list_item:
            items = []
            while index < len(lines):
                current_item = parse_list_item(lines[index])
                if not current_item:
                    break
                items.append(current_item)
                index += 1
            blocks.append({"type": "list", "items": items})
            continue
        buffer = [line.strip()]
        index += 1
        while index < len(lines):
            current = lines[index]
            current_stripped = current.strip()
            if not current_stripped:
                index += 1
                break
            if (
                current_stripped.startswith("```")
                or current_stripped == "$$"
                or current_stripped == r"\["
                or re.match(r"^(#{1,6})\s+", current)
                or re.match(r"^\s*>\s?", current)
                or parse_list_item(current) is not None
                or re.match(r"^\s*---+\s*$", current_stripped)
            ):
                break
            buffer.append(current.strip())
            index += 1
        blocks.append({"type": "paragraph", "runs": parse_inline(" ".join(buffer))})
    return blocks


def collect_math_items(blocks: list[dict]) -> list[dict[str, str]]:
    math_items: list[dict[str, str]] = []
    counter = 0

    def append_math(text: str, *, display: bool = False) -> None:
        nonlocal counter
        math_items.append({"placeholder": f"[[EQ_{counter}]]", "text": text, "display": display})
        counter += 1

    for block in blocks:
        if block["type"] in {"heading", "paragraph", "blockquote"}:
            for run in block["runs"]:
                if run["type"] == "math":
                    append_math(run["text"], display=False)
        elif block["type"] == "list":
            for item in block["items"]:
                for run in item["runs"]:
                    if run["type"] == "math":
                        append_math(run["text"], display=False)
        elif block["type"] == "math_block":
            append_math(block["text"], display=True)
    return math_items


class MarkdownWordApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("聊天 Markdown 导出 Word")
        self.root.geometry("980x720")
        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text="将聊天 Markdown 导出为 Word（保留中文排版、标题层级与常见公式）",
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        title.pack(anchor="w")

        description = ttk.Label(
            container,
            text="当前采用 python-docx 负责正文排版，Word COM 负责原生公式落地。",
        )
        description.pack(anchor="w", pady=(4, 10))

        button_bar = ttk.Frame(container)
        button_bar.pack(fill="x", pady=(0, 8))

        ttk.Button(button_bar, text="粘贴剪贴板", command=self.paste_clipboard).pack(side="left")
        ttk.Button(button_bar, text="载入 .md 文件", command=self.open_markdown_file).pack(side="left", padx=8)
        ttk.Button(button_bar, text="导出到 Word", command=self.export_to_word).pack(side="left")
        ttk.Button(button_bar, text="插入示例", command=self.insert_demo).pack(side="left", padx=8)

        self.text = tk.Text(container, wrap="word", undo=True, font=("Consolas", 11))
        self.text.pack(fill="both", expand=True)

        ttk.Label(container, text="状态").pack(anchor="w", pady=(10, 4))
        self.log = tk.Text(container, height=8, wrap="word", state="disabled", font=("Consolas", 10))
        self.log.pack(fill="x")

    def append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"{message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def paste_clipboard(self) -> None:
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showerror("读取失败", "剪贴板里没有可读取的文本内容。")
            return
        self.text.delete("1.0", "end")
        self.text.insert("1.0", text)
        self.append_log("已从剪贴板读取内容。")

    def open_markdown_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[("Markdown", "*.md *.markdown *.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        content = Path(path).read_text(encoding="utf-8")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", content)
        self.append_log(f"已载入文件：{path}")

    def insert_demo(self) -> None:
        demo = """# 示例对话

这是一个带有 **加粗**、*斜体*、`行内代码` 和行内公式 \\(x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}\\) 的段落。

## 列表

- 一元二次公式：\\(x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}\\)
  - 集合表示：\\(\\{ f_i \\in \\mathcal{F} \\mid E_s(f_i) \\ge \\tau \\}\\)
1. 求和：\\(\\sum_{i=1}^{n} x_i^2\\)
2. 积分：\\(\\int_0^1 x^2 dx = \\frac{1}{3}\\)

---

> 这是引用块，适合保留回答里的提示信息。

```python
def hello():
    return "world"
```

\\[
\\Gamma_s : 2^{\\mathcal{F}} \\rightarrow \\mathcal{G}_s
\\]
"""
        self.text.delete("1.0", "end")
        self.text.insert("1.0", demo)
        self.append_log("已插入示例内容。")

    def export_to_word(self) -> None:
        markdown_text = self.text.get("1.0", "end").strip()
        if not markdown_text:
            messagebox.showwarning("No content", "Please paste or enter Markdown content first.")
            return
        output_path = filedialog.asksaveasfilename(
            title="Save Word document",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
            initialfile="chat-export.docx",
        )
        if not output_path:
            return

        json_path: str | None = None
        layout_docx_path: str | None = None
        try:
            normalized_markdown = normalize_markdown_content(markdown_text)
            blocks = parse_markdown(normalized_markdown)
            payload = {
                "blocks": blocks,
                "math_items": collect_math_items(blocks),
            }

            with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=2)
                json_path = handle.name

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
                layout_docx_path = handle.name

            self.append_log("Building layout DOCX...")
            build_document(payload, layout_docx_path)
            self.append_log(f"Layout built: {layout_docx_path}")

            self.append_log("Injecting native OMML equations...")
            inserted_count = replace_math_placeholders(Path(layout_docx_path), Path(json_path))
            self.append_log(f"Native OMML inserted: {inserted_count}")

            shutil.copyfile(layout_docx_path, output_path)
            self.append_log(f"Export succeeded: {output_path}")
            messagebox.showinfo("Export complete", f"Document created:\n{output_path}")
        except Exception as exc:
            self.append_log(f"Export exception: {exc}")
            messagebox.showerror("Export exception", str(exc))
        finally:
            for temp_path in (json_path, layout_docx_path):
                if temp_path:
                    try:
                        Path(temp_path).unlink(missing_ok=True)
                    except OSError:
                        pass

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    MarkdownWordApp().run()
