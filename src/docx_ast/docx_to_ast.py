"""
Read a Word ``.docx`` package into an AST, and write it back out.

:class:`Document` is both the loaded package and the root of the tree: it holds
the parsed ``document.xml`` / ``footnotes.xml`` / ``comments.xml`` roots plus the
original zip, and the paragraphs below it are views of ``w:p`` elements in those
roots. Editing operations mutate that XML in place (see :mod:`xml_utils`) and
then call :meth:`Paragraph.refresh` to re-derive the AST for the paragraph they
touched.

That invariant — *XML is the source of truth, the AST is derived* — is what keeps
everything the AST does not model (bookmarks, drawings, field codes, unusual run
content, other authors' revisions) byte-identical through an edit. Saving copies
the original package entry by entry and replaces only the parts that were
touched.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree

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
    link_parents,
)
from xml_utils.element_helpers import comment_element, comments_root
from xml_utils.ns import RUN_CONTAINERS, is_element, ln, q
from xml_utils.revisions import scan_max_rev_id

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)

DOCUMENT_PART = "word/document.xml"
FOOTNOTES_PART = "word/footnotes.xml"
COMMENTS_PART = "word/comments.xml"


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone=True)


def parse_paragraph(p_el: etree._Element) -> Paragraph:
    """Parse a ``w:p`` into a Paragraph that points back at it."""
    pPr = p_el.find(q("w:pPr"))
    pStyle_el = pPr.find(q("w:pStyle")) if pPr is not None else None
    para = Paragraph(
        p_el,
        pStyle=pStyle_el.get(q("w:val")) if pStyle_el is not None else None,
    )

    raw: list[Node] = []
    for child in p_el:
        if not is_element(child):
            continue
        tag = ln(child)
        if tag == "pPr":
            continue
        if tag == "commentRangeStart":
            raw.append(CommentRangeStart(int(child.get(q("w:id"), "0"))))
        elif tag == "commentRangeEnd":
            raw.append(CommentRangeEnd(int(child.get(q("w:id"), "0"))))
        else:
            raw.extend(_parse_run_level(child))
    para.children.extend(_merge(raw))
    return para


def _parse_run_level(el: etree._Element) -> list[Node]:
    tag = ln(el)
    if tag == "r":
        return _parse_run(el)
    if tag == "hyperlink":
        node = Hyperlink(attrs=dict(el.attrib))
        node.children.extend(_parse_children(el))
        return [node]
    if tag in ("ins", "del"):
        cls = Ins if tag == "ins" else Del
        node = cls(
            id=int(el.get(q("w:id"), "0")),
            author=el.get(q("w:author"), ""),
            date=el.get(q("w:date")),
        )
        node.children.extend(_parse_children(el))
        return [node]
    if tag in RUN_CONTAINERS:  # moveFrom / moveTo
        return _parse_children(el)
    if tag == "sdt":
        content = el.find(q("w:sdtContent"))
        return _parse_children(content) if content is not None else []
    return []


def _parse_children(el: etree._Element) -> list[Node]:
    raw: list[Node] = []
    for child in el:
        if is_element(child):
            raw.extend(_parse_run_level(child))
    return _merge(raw)


def _parse_run(r_el: etree._Element) -> list[Node]:
    rPr = r_el.find(q("w:rPr"))
    highlight_el = rPr.find(q("w:highlight")) if rPr is not None else None
    highlight = highlight_el.get(q("w:val")) if highlight_el is not None else None

    out: list[Node] = []
    for child in r_el:
        if not is_element(child):
            continue
        tag = ln(child)
        if tag == "t":
            out.append(Text(child.text or "", highlight=highlight))
        elif tag == "delText":
            out.append(DelText(child.text or ""))
        elif tag == "footnoteReference":
            out.append(FootnoteReference(int(child.get(q("w:id"), "0"))))
        elif tag == "tab":
            out.append(Tab())
        # NOTE: w:commentReference is Word's clickable anchor for a comment, which
        # the range markers already locate; w:fldChar / w:instrText carry field
        # plumbing, not document text. Neither gets a node; the XML keeps both.
    return out


def _merge(nodes: list[Node]) -> list[Node]:
    """Join runs of adjacent text that the reader sees as one stretch."""
    out: list[Node] = []
    buf: Text | DelText | None = None

    for n in nodes:
        if (
            isinstance(buf, Text)
            and isinstance(n, Text)
            and n.highlight == buf.highlight
        ) or (isinstance(buf, DelText) and isinstance(n, DelText)):
            buf.text += n.text
            continue
        if isinstance(n, (Text, DelText)):
            buf = n
        else:
            buf = None
        out.append(n)
    return out


@dataclass
class Document(Node):
    """A loaded ``.docx``: the package, and the root of the AST view of it."""

    children: list[Node] = field(default_factory=list)
    footnotes: dict[int, Footnote] = field(default_factory=dict)
    comments: dict[int, Comment] = field(default_factory=dict)

    _package: bytes = field(default=b"", repr=False, compare=False)
    _doc_root: etree._Element | None = field(default=None, repr=False, compare=False)
    _fn_root: etree._Element | None = field(default=None, repr=False, compare=False)
    _cmt_root: etree._Element | None = field(default=None, repr=False, compare=False)
    _had_comments_part: bool = field(default=False, repr=False, compare=False)
    _rev_seq: int = field(default=0, repr=False, compare=False)
    _comment_seq: int = field(default=-1, repr=False, compare=False)

    @classmethod
    def load(cls, path: str | Path) -> Document:
        doc = cls(_package=Path(path).read_bytes())
        with zipfile.ZipFile(BytesIO(doc._package)) as z:
            names = set(z.namelist())
            parts = {
                name: z.read(name)
                for name in (DOCUMENT_PART, FOOTNOTES_PART, COMMENTS_PART)
                if name in names
            }

        doc._doc_root = etree.fromstring(parts[DOCUMENT_PART])
        if FOOTNOTES_PART in parts:
            doc._fn_root = etree.fromstring(parts[FOOTNOTES_PART])
        if COMMENTS_PART in parts:
            doc._cmt_root = etree.fromstring(parts[COMMENTS_PART])
            doc._had_comments_part = True

        doc._parse_body()
        doc._parse_footnotes()
        doc._parse_comments()
        doc._seed_ids()
        link_parents(doc)
        return doc

    def _parse_body(self) -> None:
        body = self._doc_root.find(q("w:body"))
        if body is None:
            return
        # A deep search, so paragraphs inside tables and text boxes are found too.
        for p_el in body.findall(f".//{q('w:p')}"):
            self.children.append(parse_paragraph(p_el))

    def _parse_footnotes(self) -> None:
        if self._fn_root is None:
            return
        for fn_el in self._fn_root.findall(q("w:footnote")):
            # w:type marks the separator / continuation pseudo-footnotes.
            if fn_el.get(q("w:type")) is not None:
                continue
            fid = int(fn_el.get(q("w:id"), "0"))
            self.footnotes[fid] = Footnote(
                id=fid,
                children=[parse_paragraph(p) for p in fn_el.findall(q("w:p"))],
            )

    def _parse_comments(self) -> None:
        if self._cmt_root is None:
            return
        for c_el in self._cmt_root.findall(q("w:comment")):
            cid = int(c_el.get(q("w:id"), "0"))
            self.comments[cid] = Comment(
                id=cid,
                author=c_el.get(q("w:author"), ""),
                date=c_el.get(q("w:date")),
                initials=c_el.get(q("w:initials"), ""),
                children=[parse_paragraph(p) for p in c_el.findall(q("w:p"))],
            )

    def _seed_ids(self) -> None:
        roots = [
            r for r in (self._doc_root, self._fn_root, self._cmt_root) if r is not None
        ]
        self._rev_seq = max(scan_max_rev_id(r) for r in roots)

        used = set(self.comments)
        for root in roots:
            for el in root.iter(q("w:commentRangeStart")):
                used.add(int(el.get(q("w:id"), "0")))
        self._comment_seq = max(used) if used else -1

    def next_rev_id(self) -> int:
        self._rev_seq += 1
        return self._rev_seq

    def next_comment_id(self) -> int:
        self._comment_seq += 1
        return self._comment_seq

    def add_comment_body(
        self,
        cid: int,
        note: str,
        *,
        author: str = "",
        date: str | None = None,
        initials: str = "",
    ) -> Comment:
        if self._cmt_root is None:
            self._cmt_root = comments_root()
        el = comment_element(cid, note, author=author, initials=initials, date=date)
        self._cmt_root.append(el)

        comment = Comment(
            id=cid,
            author=author,
            date=date,
            initials=initials,
            children=[parse_paragraph(p) for p in el.findall(q("w:p"))],
        )
        self.comments[cid] = comment
        link_parents(comment, self)
        return comment

    def save(self, out_path: str | Path) -> Path:
        replacements: dict[str, bytes] = {DOCUMENT_PART: _serialize(self._doc_root)}
        if self._fn_root is not None:
            replacements[FOOTNOTES_PART] = _serialize(self._fn_root)
        if self._cmt_root is not None:
            replacements[COMMENTS_PART] = _serialize(self._cmt_root)

        out = BytesIO()
        with (
            zipfile.ZipFile(BytesIO(self._package)) as zin,
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            names = set(zin.namelist())
            if self._cmt_root is not None and not self._had_comments_part:
                # The document had no comments part; declare the one we just made.
                replacements["[Content_Types].xml"] = _with_comments_type(
                    zin.read("[Content_Types].xml")
                )
                rels = "word/_rels/document.xml.rels"
                replacements[rels] = _with_comments_rel(
                    zin.read(rels) if rels in names else None
                )
            for item in zin.infolist():
                data = replacements.pop(item.filename, None)
                zout.writestr(item, zin.read(item.filename) if data is None else data)
            for name, data in replacements.items():
                zout.writestr(name, data)

        out_path = Path(out_path)
        out_path.write_bytes(out.getvalue())
        return out_path


def _with_comments_type(data: bytes) -> bytes:
    root = etree.fromstring(data)
    part = f"/{COMMENTS_PART}"
    for override in root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
        if override.get("PartName") == part:
            return data
    override = etree.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override")
    override.set("PartName", part)
    override.set("ContentType", COMMENTS_CONTENT_TYPE)
    return _serialize(root)


def _with_comments_rel(data: bytes | None) -> bytes:
    tag = f"{{{PACKAGE_RELS_NS}}}Relationship"
    if data is None:
        root = etree.Element(f"{{{PACKAGE_RELS_NS}}}Relationships")
    else:
        root = etree.fromstring(data)
        for rel in root.findall(tag):
            if rel.get("Type") == COMMENTS_REL_TYPE:
                return data
    used = {rel.get("Id") for rel in root.findall(tag)}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rel = etree.SubElement(root, tag)
    rel.set("Id", f"rId{n}")
    rel.set("Type", COMMENTS_REL_TYPE)
    rel.set("Target", "comments.xml")
    return _serialize(root)
