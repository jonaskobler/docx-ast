from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from lxml import etree

_WS_VARIANTS = (
    "\u00a0\u2002\u2003\u2004\u2005\u2006\u2007"
    "\u2008\u2009\u200a\u202f\u205f\u3000"
)
_WS_TABLE = {ord(c): " " for c in _WS_VARIANTS}


def normalize_ws(s: str) -> str:
    return s.translate(_WS_TABLE)


@dataclass
class RunStyle:
    """Inline (run-level) formatting properties."""

    rStyle: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    font_name: str | None = None
    font_size: int | None = None
    color: str | None = None
    highlight: str | None = None

    _raw: etree.Element | None = field(
        default=None, init=False, repr=False, compare=False
    )

    @classmethod
    def from_xml(cls, rPr: etree.Element) -> "RunStyle":
        """Parse a <w:rPr> lxml element into a RunStyle."""
        from docx_ast.constants import q
        import copy

        def find(tag):
            return rPr.find(q(tag))

        def bool_prop(tag) -> bool | None:
            el = find(tag)
            if el is None:
                return None
            val = el.get(q("w:val"), "true")
            return val not in ("0", "false")

        rStyle_el = find("w:rStyle")
        fonts_el = find("w:rFonts")
        sz_el = find("w:sz")
        color_el = find("w:color")
        highlight_el = find("w:highlight")

        style = cls(
            rStyle=rStyle_el.get(q("w:val")) if rStyle_el is not None else None,
            bold=bool_prop("w:b"),
            italic=bool_prop("w:i"),
            font_name=(
                (fonts_el.get(q("w:ascii")) or fonts_el.get(q("w:hAnsi")))
                if fonts_el is not None
                else None
            ),
            font_size=int(sz_el.get(q("w:val"))) if sz_el is not None else None,
            color=color_el.get(q("w:val")) if color_el is not None else None,
            highlight=highlight_el.get(q("w:val"))
            if highlight_el is not None
            else None,
        )
        style._raw = copy.deepcopy(rPr)
        return style

    def with_highlight(self, color: str = "yellow") -> RunStyle:
        from docx_ast.constants import q
        from lxml import etree
        import copy

        raw = (
            copy.deepcopy(self._raw)
            if self._raw is not None
            else etree.Element(q("w:rPr"))
        )
        el = raw.find(q("w:highlight"))
        if el is None:
            el = etree.SubElement(raw, q("w:highlight"))
        el.set(q("w:val"), color)
        new = RunStyle(
            rStyle=self.rStyle,
            bold=self.bold,
            italic=self.italic,
            font_name=self.font_name,
            font_size=self.font_size,
            color=self.color,
            highlight=color,
        )
        new._raw = raw
        return new


@dataclass
class ParagraphStyle:
    """Block (paragraph-level) formatting properties."""

    pStyle: str | None = None  # named w:pStyle (e.g. "Normal", "Heading1")
    alignment: str | None = None  # "left" | "center" | "right" | "both"

    _raw: etree.Element | None = field(
        default=None, init=False, repr=False, compare=False
    )

    @classmethod
    def from_xml(cls, pPr: etree.Element) -> ParagraphStyle:
        """Parse a <w:pPr> lxml element into a ParagraphStyle."""
        from docx_ast.constants import q
        import copy

        pStyle_el = pPr.find(q("w:pStyle"))
        jc_el = pPr.find(q("w:jc"))

        style = cls(
            pStyle=pStyle_el.get(q("w:val")) if pStyle_el is not None else None,
            alignment=jc_el.get(q("w:val")) if jc_el is not None else None,
        )
        style._raw = copy.deepcopy(pPr)
        return style


@dataclass
class Node:
    """Base class for all AST nodes."""

    parent: Node | None = field(default=None, init=False, repr=False, compare=False)

    def to_str(self) -> str:
        from docx_ast.ast_to_str import ASTStringSerializer

        return ASTStringSerializer().to_str(self)


@dataclass
class Text(Node):
    """Plain textual content."""

    text: str
    style_id: str | None = None

    def alter_tracked(
        self, new_text: str, *, author: str = "", date: str | None = None
    ) -> list[Node]:
        from docx_ast.tracked_changes import alter_text_tracked

        return alter_text_tracked(self, new_text, author=author, date=date)


@dataclass
class DelText(Node):
    """Deleted textual content (w:delText), inside a Del range."""

    text: str
    style_id: str | None = None


@dataclass
class FootnoteReference(Node):
    """Inline reference to a footnote."""

    id: int
    style_id: str | None = None


@dataclass
class Tab(Node):
    style_id: str | None = None


@dataclass
class FieldChar(Node):
    kind: str = ""
    style_id: str | None = None


@dataclass
class InstrText(Node):
    text: str = ""
    style_id: str | None = None


@dataclass
class Ins(Node):
    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)


@dataclass
class Del(Node):
    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)


@dataclass
class MoveFrom(Node):
    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)
    attrs: dict = field(default_factory=dict)


@dataclass
class MoveTo(Node):
    id: int = 0
    author: str = ""
    date: str | None = None
    children: list[Node] = field(default_factory=list)
    attrs: dict = field(default_factory=dict)


@dataclass
class Hyperlink(Node):
    children: list[Node] = field(default_factory=list)
    attrs: dict = field(default_factory=dict)


@dataclass
class Sdt(Node):
    children: list[Node] = field(default_factory=list)
    pr: list = field(default_factory=list, repr=False, compare=False)


@dataclass
class Paragraph(Node):
    children: list[Node] = field(default_factory=list)
    style_id: str | None = None
    _xml: object = field(default=None, init=False, repr=False, compare=False)

    def alter_tracked(
        self, new_text: str, *, author: str = "Claude", date: str | None = None
    ) -> list[Node]:
        from docx_ast.tracked_changes import alter_paragraph_tracked

        return alter_paragraph_tracked(self, new_text, author=author, date=date)

    def get_text(self) -> str:
        return "".join(n.text for n in walk(self) if isinstance(n, Text))

    def highlight(self, substring: str, color: str = "yellow") -> "list[Node]":
        if not substring:
            return self.children
        nodes = [n for n in walk(self) if isinstance(n, Text)]
        hay = normalize_ws("".join(n.text for n in nodes))
        needle = normalize_ws(substring)
        spans: list[tuple[int, int]] = []
        i = hay.find(needle)
        while i >= 0:
            spans.append((i, i + len(needle)))
            i = hay.find(needle, i + len(needle))
        return self.highlight_spans(spans, color=color)

    def highlight_spans(
        self, spans: "list[tuple[int, int]]", color: str = "yellow"
    ) -> "list[Node]":
        spans = [(a, b) for a, b in spans if b > a]
        if not spans:
            return self.children
        doc = self
        while doc is not None and not isinstance(doc, Document):
            doc = doc.parent
        if doc is None:
            raise ValueError("paragraph is not attached to a Document")
        nodes = [n for n in walk(self) if isinstance(n, Text)]

        def covered(a: int, b: int) -> bool:
            return any(s <= a and b <= e for s, e in spans)

        pos = 0
        for n in nodes:
            ns, ne = pos, pos + len(n.text)
            pos = ne
            if not any(s < ne and e > ns for s, e in spans):
                continue
            points = {ns, ne}
            for s, e in spans:
                if ns < e and ne > s:
                    points.add(min(max(s, ns), ne))
                    points.add(min(max(e, ns), ne))
            ordered = sorted(points)
            replacement: list[Node] = []
            for k in range(len(ordered) - 1):
                a, b = ordered[k], ordered[k + 1]
                if b <= a:
                    continue
                seg = n.text[a - ns : b - ns]
                if covered(a, b):
                    base = doc.styles.get(n.style_id) if n.style_id else None
                    if not isinstance(base, RunStyle):
                        base = RunStyle()
                    hl_id = doc.register_style(base.with_highlight(color))
                    replacement.append(Text(seg, style_id=hl_id))
                else:
                    replacement.append(Text(seg, style_id=n.style_id))
            kids = n.parent.children
            idx = next(j for j, c in enumerate(kids) if c is n)
            kids[idx : idx + 1] = replacement
        link_parents(self, self.parent)
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
        if not substring:
            return None
        doc = self
        while doc is not None and not isinstance(doc, Document):
            doc = doc.parent
        if doc is None:
            raise ValueError("paragraph is not attached to a Document")
        nodes = [n for n in walk(self) if isinstance(n, Text)]
        full = "".join(n.text for n in nodes)
        start = normalize_ws(full).find(normalize_ws(substring))
        if start < 0:
            return None
        end = start + len(substring)

        used = set(doc.comments)
        for n in walk(doc):
            if isinstance(n, (CommentRangeStart, CommentRangeEnd)):
                used.add(n.id)
        cid = max(used) + 1 if used else 0
        start_node = CommentRangeStart(cid)
        end_node = CommentRangeEnd(cid)

        opened = closed = False
        pos = 0
        for n in nodes:
            s, e = pos, pos + len(n.text)
            pos = e
            cuts: list[tuple[int, Node]] = []
            if not opened and s <= start <= e:
                cuts.append((start - s, start_node))
                opened = True
            if not closed and s <= end <= e:
                cuts.append((end - s, end_node))
                closed = True
            if not cuts:
                continue
            cuts.sort(key=lambda c: c[0])
            replacement: list[Node] = []
            last = 0
            for off, marker in cuts:
                if off > last:
                    replacement.append(Text(n.text[last:off], style_id=n.style_id))
                replacement.append(marker)
                last = off
            if last < len(n.text):
                replacement.append(Text(n.text[last:], style_id=n.style_id))
            kids = n.parent.children
            idx = next(i for i, c in enumerate(kids) if c is n)
            kids[idx : idx + 1] = replacement

        doc.comments[cid] = Comment(
            id=cid,
            author=author,
            date=date,
            initials=initials,
            content=[Paragraph(children=[Text(note)])],
        )
        link_parents(doc)
        return cid


@dataclass
class PageBreak(Node):
    """A page boundary. page is the number of the page that begins here, or None."""

    page: int | None = None


@dataclass
class Footnote(Node):
    """Footnote content, keyed by id and emitted into footnotes.xml."""

    id: int
    children: list[Node] = field(default_factory=list)
    style_id: str | None = None


@dataclass
class Comment(Node):
    """Comment content, keyed by id and emitted into comments.xml."""

    id: int
    author: str = ""
    date: str | None = None
    initials: str = ""
    content: list[Node] = field(default_factory=list)


@dataclass
class CommentRangeStart(Node):
    """Inline marker where a comment range opens."""

    id: int


@dataclass
class CommentRangeEnd(Node):
    """Inline marker where a comment range closes."""

    id: int


@dataclass
class MoveRangeMarker(Node):
    kind: str = ""
    attrs: dict = field(default_factory=dict)


@dataclass
class Document(Node):
    """Root node."""

    children: list[Node] = field(default_factory=list)
    footnotes: dict[int, "Footnote"] = field(default_factory=dict)
    comments: dict[int, "Comment"] = field(default_factory=dict)
    styles: dict[str, "RunStyle | ParagraphStyle"] = field(default_factory=dict)
    _style_keys: dict = field(
        default_factory=dict, init=False, repr=False, compare=False
    )
    _run_seq: int = field(default=0, init=False, repr=False, compare=False)
    _para_seq: int = field(default=0, init=False, repr=False, compare=False)

    def register_style(self, style: "RunStyle | ParagraphStyle") -> str:

        raw = getattr(style, "_raw", None)
        key = etree.tostring(raw, encoding="ascii") if raw is not None else None
        # NOTE: apparently the style_id in docx is always ASCII, all non ascii chars get dropped,
        # the proper name is stored in the styles.xml und w:name
        if key is not None and key in self._style_keys:
            return self._style_keys[key]
        if isinstance(style, RunStyle):
            sid = f"r{self._run_seq}"
            self._run_seq += 1
        else:
            sid = f"p{self._para_seq}"
            self._para_seq += 1
        self.styles[sid] = style
        if key is not None:
            self._style_keys[key] = sid
        return sid

    def comment_text(self, cid: int) -> str:
        out: list[str] = []
        active = False
        for node in walk(self):
            if isinstance(node, CommentRangeStart) and node.id == cid:
                active = True
            elif isinstance(node, CommentRangeEnd) and node.id == cid:
                active = False
            elif active and isinstance(node, Text):
                out.append(node.text)
        return "".join(out)


def walk(node: Node) -> Iterator[Node]:
    yield node
    for child in getattr(node, "children", None) or []:
        yield from walk(child)


def link_parents(node: Node, parent: Node | None = None) -> None:
    node.parent = parent
    for child in getattr(node, "children", None) or []:
        link_parents(child, node)
    for note in getattr(node, "content", None) or []:
        link_parents(note, node)
    footnotes = getattr(node, "footnotes", None)
    if footnotes:
        for fn in footnotes.values():
            link_parents(fn, node)
    comments = getattr(node, "comments", None)
    if comments:
        for comment in comments.values():
            link_parents(comment, node)
