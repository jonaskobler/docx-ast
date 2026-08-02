from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field

from lxml import etree

from xml_utils.xml_edit import ParagraphIndex

_WS_VARIANTS = (
    "\u00a0\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u202f\u205f\u3000"
)
_WS_TABLE = {ord(c): " " for c in _WS_VARIANTS}


def normalize_ws(s: str) -> str:
    return s.translate(_WS_TABLE)


@dataclass
class Node:
    """Base class for all AST nodes."""

    parent: Node | None = field(default=None, init=False, repr=False, compare=False)

    @property
    def root(self) -> Node:
        node = self
        while node.parent is not None:
            node = node.parent
        return node

    def to_str(self) -> str:
        from docx_ast.ast_to_str import to_str

        return to_str(self)


@dataclass
class Text(Node):
    """Visible textual content (``w:t``)."""

    text: str
    # w:rPr/w:highlight, the one run property the AST models: highlighting is an
    # operation this library offers, so the view has to be able to show one.
    highlight: str | None = None

    def alter_tracked(
        self, new_text: str, *, author: str = "", date: str | None = None
    ) -> list[Node]:
        from docx_ast.tracked_changes import alter_text_tracked

        return alter_text_tracked(self, new_text, author=author, date=date)


@dataclass
class DelText(Node):
    """Deleted textual content (``w:delText``), inside a :class:`Del`.

    Not part of the paragraph's text: it is no longer visible, and neither
    ``get_text()`` nor the offsets in :mod:`xml_utils.xml_edit` count it.
    """

    text: str


@dataclass
class FootnoteReference(Node):
    """Inline reference to a footnote."""

    id: int


@dataclass
class Tab(Node):
    """A ``w:tab``. Contributes no characters to the paragraph's text."""


@dataclass
class Ins(Node):
    """A tracked insertion (``w:ins``)."""

    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)


@dataclass
class Del(Node):
    """A tracked deletion (``w:del``)."""

    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)


@dataclass
class Hyperlink(Node):
    children: list[Node] = field(default_factory=list)
    attrs: dict = field(default_factory=dict)


@dataclass
class CommentRangeStart(Node):
    """Inline marker where a comment range opens."""

    id: int


@dataclass
class CommentRangeEnd(Node):
    """Inline marker where a comment range closes."""

    id: int


@dataclass
class Paragraph(Node):
    """A view of one ``w:p``.

    ``_xml`` is required: the XML is the source of truth and the children are only
    ever re-derived from it, so a paragraph without one could not be edited,
    indexed or refreshed.
    """

    _xml: etree._Element = field(repr=False, compare=False)
    children: list[Node] = field(default_factory=list)
    pStyle: str | None = None  # w:pPr/w:pStyle, e.g. "Heading1", "FootnoteText"

    def get_text(self) -> str:
        return ParagraphIndex(self._xml).text

    def refresh(self) -> Paragraph:
        from docx_ast.docx_to_ast import parse_paragraph

        fresh = parse_paragraph(self._xml)
        self.children[:] = fresh.children
        self.pStyle = fresh.pStyle
        link_parents(self, self.parent)
        return self

    def alter_tracked(
        self, new_text: str, *, author: str = "", date: str | None = None
    ) -> list[Node]:
        from docx_ast.tracked_changes import alter_paragraph_tracked

        return alter_paragraph_tracked(self, new_text, author=author, date=date)

    def highlight(self, substring: str, color: str = "yellow") -> list[Node]:
        if not substring:
            return self.children
        hay = normalize_ws(self.get_text())
        needle = normalize_ws(substring)
        spans: list[tuple[int, int]] = []
        i = hay.find(needle)
        while i >= 0:
            spans.append((i, i + len(needle)))
            i = hay.find(needle, i + len(needle))
        return self.highlight_spans(spans, color=color)

    def highlight_spans(
        self, spans: list[tuple[int, int]], color: str = "yellow"
    ) -> list[Node]:
        """Highlight character ranges given as offsets into ``get_text()``."""
        from xml_utils.xml_edit import highlight_spans_xml

        spans = [(a, b) for a, b in spans if b > a]
        if not spans:
            return self.children
        highlight_spans_xml(ParagraphIndex(self._xml), spans, color)
        self.refresh()
        return self.children

    def add_comment(
        self,
        substring: str,
        note: str,
        *,
        author: str = "",
        date: str | None = None,
        initials: str = "",
    ) -> int | None:
        """Attach a Word comment to the first occurrence of ``substring``.

        Returns the new comment id, or None when the substring is not in this
        paragraph. A span that cannot be anchored raises, rather than failing
        quietly.
        """
        from docx_ast.docx_to_ast import Document
        from xml_utils.xml_edit import add_comment_range_xml

        if not substring:
            return None
        doc = self.root
        if not isinstance(doc, Document):
            raise ValueError("paragraph is not attached to a Document")

        start = normalize_ws(self.get_text()).find(normalize_ws(substring))
        if start < 0:
            return None
        cid = doc.next_comment_id()
        add_comment_range_xml(
            ParagraphIndex(self._xml), start, start + len(substring), cid
        )
        doc.add_comment_body(cid, note, author=author, date=date, initials=initials)
        self.refresh()
        return cid


@dataclass
class Footnote(Node):
    """A footnote's paragraphs, keyed by id in ``Document.footnotes``."""

    id: int
    children: list[Node] = field(default_factory=list)


@dataclass
class Comment(Node):
    """A comment's body paragraphs, keyed by id in ``Document.comments``."""

    id: int
    author: str = ""
    date: str | None = None
    initials: str = ""
    children: list[Node] = field(default_factory=list)


def walk(node: Node) -> Iterator[Node]:
    yield node
    for child in getattr(node, "children", None) or []:
        yield from walk(child)


def link_parents(node: Node, parent: Node | None = None) -> None:
    node.parent = parent
    for child in getattr(node, "children", None) or []:
        link_parents(child, node)
    for part in (getattr(node, "footnotes", None), getattr(node, "comments", None)):
        for entry in (part or {}).values():
            link_parents(entry, node)
