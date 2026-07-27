"""
Pretty-print an AST as an indented tree, for debugging.

Mirrors ``ASTDocxSerializer``: a ``functools.singledispatchmethod`` (`_to_str`)
handles each node type and recurses into the children one indent level deeper.

    from docx_ast.ast_to_str import ASTStringSerializer
    print(ASTStringSerializer().to_str(document))
"""

from __future__ import annotations

from functools import singledispatchmethod

from docx_ast.nodes import (
    Comment,
    CommentRangeEnd,
    CommentRangeStart,
    Del,
    DelText,
    Document,
    FieldChar,
    Footnote,
    FootnoteReference,
    Hyperlink,
    Ins,
    InstrText,
    MoveFrom,
    MoveRangeMarker,
    MoveTo,
    Node,
    PageBreak,
    Paragraph,
    Sdt,
    Tab,
    Text,
)

INDENT = "  "
_TEXT_MAX = 80


def _short(text: str) -> str:
    text = " ".join(text.split())
    return text if len(text) <= _TEXT_MAX else text[: _TEXT_MAX - 1] + "…"


class ASTStringSerializer:
    """Render an AST node (and its descendants) as an indented tree string."""

    def to_str(self, node: Node) -> str:
        return self._to_str(node, 0)

    # ── dispatch ───────────────────────────────────────────────────────────────
    @singledispatchmethod
    def _to_str(self, node: Node, depth: int) -> str:
        return f"{INDENT * depth}{type(node).__name__}"

    def _block(self, label: str, children: list[Node], depth: int) -> str:
        lines = [f"{INDENT * depth}{label}"]
        lines += [self._to_str(child, depth + 1) for child in children]
        return "\n".join(lines)

    @_to_str.register(Document)
    def _(self, node: Document, depth: int) -> str:
        lines = [self._block("Document", node.children, depth)]
        if node.footnotes:
            lines.append(f"{INDENT * (depth + 1)}footnotes:")
            for fid in sorted(node.footnotes):
                lines.append(self._to_str(node.footnotes[fid], depth + 2))
        if node.comments:
            lines.append(f"{INDENT * (depth + 1)}comments:")
            for cid in sorted(node.comments):
                lines.append(self._to_str(node.comments[cid], depth + 2))
        return "\n".join(lines)

    @_to_str.register(Paragraph)
    def _(self, node: Paragraph, depth: int) -> str:
        return self._block("Paragraph", node.children, depth)

    @_to_str.register(Footnote)
    def _(self, node: Footnote, depth: int) -> str:
        return self._block(f"Footnote(id={node.id})", node.children, depth)

    @_to_str.register(Comment)
    def _(self, node: Comment, depth: int) -> str:
        label = f"Comment(id={node.id}, author={node.author!r})"
        return self._block(label, node.content, depth)

    @_to_str.register(CommentRangeStart)
    def _(self, node: CommentRangeStart, depth: int) -> str:
        return f"{INDENT * depth}CommentRangeStart(id={node.id})"

    @_to_str.register(CommentRangeEnd)
    def _(self, node: CommentRangeEnd, depth: int) -> str:
        return f"{INDENT * depth}CommentRangeEnd(id={node.id})"

    @_to_str.register(Text)
    def _(self, node: Text, depth: int) -> str:
        return f"{INDENT * depth}Text({_short(node.text)!r})"

    @_to_str.register(DelText)
    def _(self, node: DelText, depth: int) -> str:
        return f"{INDENT * depth}DelText({_short(node.text)!r})"

    @_to_str.register(Ins)
    def _(self, node: Ins, depth: int) -> str:
        return self._block(self._rev_label("Ins", node), node.children, depth)

    @_to_str.register(Del)
    def _(self, node: Del, depth: int) -> str:
        return self._block(self._rev_label("Del", node), node.children, depth)

    @staticmethod
    def _rev_label(name: str, node: Ins | Del) -> str:
        meta = f"id={node.id}, author={node.author!r}"
        if node.date is not None:
            meta += f", date={node.date!r}"
        return f"{name}({meta})"

    @_to_str.register(FootnoteReference)
    def _(self, node: FootnoteReference, depth: int) -> str:
        return f"{INDENT * depth}FootnoteReference(id={node.id})"

    @_to_str.register(PageBreak)
    def _(self, node: PageBreak, depth: int) -> str:
        return f"{INDENT * depth}PageBreak(page={node.page})"

    @_to_str.register(MoveFrom)
    def _(self, node: MoveFrom, depth: int) -> str:
        return self._block(self._rev_label("MoveFrom", node), node.children, depth)

    @_to_str.register(MoveTo)
    def _(self, node: MoveTo, depth: int) -> str:
        return self._block(self._rev_label("MoveTo", node), node.children, depth)

    @_to_str.register(Hyperlink)
    def _(self, node: Hyperlink, depth: int) -> str:
        return self._block("Hyperlink", node.children, depth)

    @_to_str.register(Sdt)
    def _(self, node: Sdt, depth: int) -> str:
        return self._block("Sdt", node.children, depth)

    @_to_str.register(Tab)
    def _(self, node: Tab, depth: int) -> str:
        return f"{INDENT * depth}Tab"

    @_to_str.register(FieldChar)
    def _(self, node: FieldChar, depth: int) -> str:
        return f"{INDENT * depth}FieldChar({node.kind!r})"

    @_to_str.register(InstrText)
    def _(self, node: InstrText, depth: int) -> str:
        return f"{INDENT * depth}InstrText({_short(node.text)!r})"

    @_to_str.register(MoveRangeMarker)
    def _(self, node: MoveRangeMarker, depth: int) -> str:
        return f"{INDENT * depth}MoveRangeMarker({node.kind!r})"


if __name__ == "__main__":
    from docx_ast.nodes import (
        Document,
        Footnote,
        FootnoteReference,
        Paragraph,
        Text,
    )

    doc = Document(
        children=[
            Paragraph(children=[Text("Olaf Sosnitza"), FootnoteReference(1)]),
        ],
        footnotes={
            1: Footnote(id=1, children=[Text("Prof. Dr. jur., Würzburg.")]),
        },
    )
    print(ASTStringSerializer().to_str(doc))
