"""
Pretty-print an AST as an indented tree, for debugging.

    from docx_ast import Document
    print(Document.load("input.docx").to_str())
"""

from __future__ import annotations

from docx_ast.docx_to_ast import Document
from docx_ast.nodes import (
    Comment,
    CommentRangeEnd,
    CommentRangeStart,
    Del,
    DelText,
    Footnote,
    FootnoteReference,
    Hyperlink,
    Ins,
    Node,
    Paragraph,
    Tab,
    Text,
)

INDENT = "  "
_TEXT_MAX = 80


def _short(text: str) -> str:
    text = " ".join(text.split())
    return text if len(text) <= _TEXT_MAX else text[: _TEXT_MAX - 1] + "…"


def _block(label: str, children: list[Node], depth: int) -> str:
    lines = [f"{INDENT * depth}{label}"]
    lines += [to_str(child, depth + 1) for child in children]
    return "\n".join(lines)


def _rev_label(name: str, node: Ins | Del) -> str:
    meta = f"id={node.id}, author={node.author!r}"
    if node.date is not None:
        meta += f", date={node.date!r}"
    return f"{name}({meta})"


def to_str(node: Node, depth: int = 0) -> str:
    """Render ``node`` and its descendants as an indented tree."""
    pad = INDENT * depth

    if isinstance(node, Text):
        hl = f", highlight={node.highlight!r}" if node.highlight else ""
        return f"{pad}Text({_short(node.text)!r}{hl})"
    if isinstance(node, DelText):
        return f"{pad}DelText({_short(node.text)!r})"
    if isinstance(node, Tab):
        return f"{pad}Tab"
    if isinstance(node, FootnoteReference):
        return f"{pad}FootnoteReference(id={node.id})"
    if isinstance(node, CommentRangeStart):
        return f"{pad}CommentRangeStart(id={node.id})"
    if isinstance(node, CommentRangeEnd):
        return f"{pad}CommentRangeEnd(id={node.id})"
    if isinstance(node, Ins):
        return _block(_rev_label("Ins", node), node.children, depth)
    if isinstance(node, Del):
        return _block(_rev_label("Del", node), node.children, depth)
    if isinstance(node, Hyperlink):
        return _block("Hyperlink", node.children, depth)
    if isinstance(node, Paragraph):
        label = "Paragraph" if node.pStyle is None else f"Paragraph({node.pStyle})"
        return _block(label, node.children, depth)
    if isinstance(node, Footnote):
        return _block(f"Footnote(id={node.id})", node.children, depth)
    if isinstance(node, Comment):
        label = f"Comment(id={node.id}, author={node.author!r})"
        return _block(label, node.children, depth)

    if isinstance(node, Document):
        lines = [_block("Document", node.children, depth)]
        for name, part in (("footnotes", node.footnotes), ("comments", node.comments)):
            if not part:
                continue
            lines.append(f"{INDENT * (depth + 1)}{name}:")
            lines += [to_str(part[k], depth + 2) for k in sorted(part)]
        return "\n".join(lines)

    return f"{pad}{type(node).__name__}"
