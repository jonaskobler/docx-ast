from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

from docx_ast.ast_to_docx import ASTDocxSerializer
from docx_ast.constants import (
    COMMENTS_CONTENT_TYPE,
    COMMENTS_REL_TYPE,
    CONTENT_TYPES_NS,
    PACKAGE_RELS_NS,
    q,
)
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
    Paragraph,
    ParagraphStyle,
    RunStyle,
    Sdt,
    Tab,
    Text,
    link_parents,
)


def _ln(el: etree._Element) -> str:
    return etree.QName(el).localname


def _xml(root: etree._Element) -> bytes:
    return etree.tostring(
        root, xml_declaration=True, encoding="utf-8", standalone=True, pretty_print=True
    )


class _CRef:
    __slots__ = ("cid",)

    def __init__(self, cid: int) -> None:
        self.cid = cid


class DocxASTExtractor:
    def __init__(self) -> None:
        self._package_bytes: bytes = b""
        self._root: etree._Element | None = None
        self._fn_root: etree._Element | None = None
        self._cmt_root: etree._Element | None = None
        self._ast: Document | None = None
        self._linked: list[Paragraph] = []

    def load(self, path: str | Path, pop: bool = False) -> Document:
        self._package_bytes = Path(path).read_bytes()
        with zipfile.ZipFile(BytesIO(self._package_bytes)) as z:
            names = set(z.namelist())
            xml = z.read("word/document.xml")
            fn_xml = (
                z.read("word/footnotes.xml") if "word/footnotes.xml" in names else None
            )
            cmt_xml = (
                z.read("word/comments.xml") if "word/comments.xml" in names else None
            )
            styles_xml = (
                z.read("word/styles.xml") if "word/styles.xml" in names else None
            )
        self._root = etree.fromstring(xml)
        self._ast = Document()
        if cmt_xml is not None:
            self._cmt_root = etree.fromstring(cmt_xml)
            self._extract_comments(self._cmt_root)
        self._extract(self._root, pop)
        if fn_xml is not None:
            self._fn_root = etree.fromstring(fn_xml)
            self._extract_footnotes(self._fn_root, pop, self._ast)
        link_parents(self._ast)
        return self._ast

    def _extract(self, root: etree._Element, pop: bool) -> Document:
        body = root.find(q("w:body"))
        if body is None:
            return self._ast
        for p_el in body.findall(f".//{q('w:p')}"):
            # HACK: We just do a deep search to avoid parsing tables and stuff for now
            self._ast.children.append(self._extract_paragraph(p_el, pop))
        return self._ast

    def _extract_comments(self, root: etree._Element) -> None:
        for c_el in root.findall(q("w:comment")):
            cid = int(c_el.get(q("w:id"), "0"))
            content = [
                self._extract_paragraph(p_el, False, link=False)
                for p_el in c_el.findall(q("w:p"))
            ]
            self._ast.comments[cid] = Comment(
                id=cid,
                author=c_el.get(q("w:author"), ""),
                date=c_el.get(q("w:date")),
                initials=c_el.get(q("w:initials"), ""),
                content=content,
            )

    def _extract_footnotes(
        self, root: etree._Element, pop: bool, doc: Document
    ) -> None:
        for fn_el in root.findall(q("w:footnote")):
            if fn_el.get(q("w:type")) is not None:
                continue
            fid = int(fn_el.get(q("w:id"), "0"))
            fn = Footnote(id=fid)
            for p_el in fn_el.findall(q("w:p")):
                fn.children.append(self._extract_paragraph(p_el, pop))
            doc.footnotes[fid] = fn

    def _extract_paragraph(
        self, p_el: etree._Element, pop: bool, link: bool = True
    ) -> Paragraph:
        style_id = None
        pPr = p_el.find(q("w:pPr"))
        if pPr is not None:
            style_id = self._register(pPr, "p")
        para = Paragraph(style_id=style_id)
        para._xml = p_el
        if link:
            self._linked.append(para)
        raw: list[Node] = []
        consumed: list[etree._Element] = []

        for child in p_el:
            tag = _ln(child)
            if tag == "pPr":
                continue
            if tag == "commentRangeStart":
                raw.append(CommentRangeStart(int(child.get(q("w:id"), "0"))))
                consumed.append(child)
                continue
            if tag == "commentRangeEnd":
                raw.append(CommentRangeEnd(int(child.get(q("w:id"), "0"))))
                consumed.append(child)
                continue
            if tag in (
                "moveFromRangeStart",
                "moveFromRangeEnd",
                "moveToRangeStart",
                "moveToRangeEnd",
            ):
                raw.append(MoveRangeMarker(tag, dict(child.attrib)))
                consumed.append(child)
                continue
            nodes = self._extract_run_level(child)
            if not nodes:
                continue
            raw.extend(nodes)
            consumed.append(child)
        merged = self._merge(raw)
        para.children.extend(n for n in merged if not isinstance(n, _CRef))
        if pop:
            for child in consumed:
                p_el.remove(child)
        return para

    def _merge(self, nodes: list[Node]) -> list[Node]:
        out: list[Node] = []
        text_buf = ""
        buf_cls: type | None = None
        buf_style: str | None = None

        def flush() -> None:
            nonlocal text_buf, buf_cls, buf_style
            if buf_cls is not None:
                out.append(buf_cls(text_buf, style_id=buf_style))
            text_buf = ""
            buf_cls = None
            buf_style = None

        for n in nodes:
            if isinstance(n, (Text, DelText)):
                if type(n) is buf_cls and n.style_id == buf_style:
                    text_buf += n.text
                else:
                    flush()
                    buf_cls = type(n)
                    buf_style = n.style_id
                    text_buf = n.text
            else:
                flush()
                out.append(n)
        flush()
        return out

    def _extract_run_level(self, el: etree._Element) -> list[Node]:
        tag = _ln(el)
        if tag == "r":
            return self._extract_run(el)
        if tag == "ins":
            return [self._extract_revision(el, Ins)]
        if tag == "del":
            return [self._extract_revision(el, Del)]
        if tag == "moveFrom":
            return [self._extract_revision(el, MoveFrom)]
        if tag == "moveTo":
            return [self._extract_revision(el, MoveTo)]
        if tag == "hyperlink":
            return [self._extract_hyperlink(el)]
        if tag == "sdt":
            return [self._extract_sdt(el)]
        return []

    def _extract_revision(self, el: etree._Element, cls) -> Node:
        node = cls(
            id=int(el.get(q("w:id"), "0")),
            author=el.get(q("w:author"), ""),
            date=el.get(q("w:date")),
        )
        if hasattr(node, "attrs"):
            node.attrs = dict(el.attrib)
        raw: list[Node] = []
        for child in el:
            raw.extend(self._extract_run_level(child))
        node.children.extend(self._merge(raw))
        return node

    def _extract_hyperlink(self, el: etree._Element) -> Hyperlink:
        node = Hyperlink(attrs=dict(el.attrib))
        raw: list[Node] = []
        for child in el:
            raw.extend(self._extract_run_level(child))
        node.children.extend(self._merge(raw))
        return node

    def _extract_sdt(self, el: etree._Element) -> Sdt:
        import copy

        node = Sdt(pr=[copy.deepcopy(c) for c in el if _ln(c) != "sdtContent"])
        content = el.find(q("w:sdtContent"))
        raw: list[Node] = []
        if content is not None:
            for child in content:
                raw.extend(self._extract_run_level(child))
        node.children.extend(self._merge(raw))
        return node

    def _extract_run(self, r_el: etree._Element) -> list[Node]:
        style_id = None
        rPr = r_el.find(q("w:rPr"))
        if rPr is not None:
            style_id = self._register(rPr, "r")
        out: list[Node] = []
        for child in r_el:
            tag = _ln(child)
            if tag == "t":
                out.append(Text(child.text or "", style_id=style_id))
            elif tag == "delText":
                out.append(DelText(child.text or "", style_id=style_id))
            elif tag == "footnoteReference":
                out.append(
                    FootnoteReference(int(child.get(q("w:id"), "0")), style_id=style_id)
                )
            elif tag == "commentReference":
                out.append(_CRef(int(child.get(q("w:id"), "0"))))
            elif tag == "tab":
                out.append(Tab(style_id=style_id))
            elif tag == "fldChar":
                out.append(
                    FieldChar(child.get(q("w:fldCharType"), ""), style_id=style_id)
                )
            elif tag == "instrText":
                out.append(InstrText(child.text or "", style_id=style_id))
        return out

    def _register(self, el: etree._Element, kind: str) -> str:
        style = RunStyle.from_xml(el) if kind == "r" else ParagraphStyle.from_xml(el)
        return self._ast.register_style(style)

    def repopulate(self) -> None:
        serializer = ASTDocxSerializer()
        serializer._styles = self._ast.styles
        for para in self._linked:
            if para._xml is None:
                continue
            for child in para.children:
                serializer._emit(child, para._xml)

    def save(self, out_path: str | Path) -> Path:
        replacements: dict[str, bytes] = {"word/document.xml": _xml(self._root)}
        if self._fn_root is not None:
            replacements["word/footnotes.xml"] = _xml(self._fn_root)

        comments = [self._ast.comments[k] for k in sorted(self._ast.comments)]
        out = BytesIO()
        with (
            zipfile.ZipFile(BytesIO(self._package_bytes)) as zin,
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            names = set(zin.namelist())
            if comments:
                serializer = ASTDocxSerializer()
                serializer._styles = self._ast.styles
                replacements["word/comments.xml"] = _xml(
                    serializer.comments_xml(comments)
                )
                if "word/comments.xml" not in names:
                    replacements["[Content_Types].xml"] = self._with_comments_type(
                        zin.read("[Content_Types].xml")
                    )
                    rels = "word/_rels/document.xml.rels"
                    replacements[rels] = self._with_comments_rel(
                        zin.read(rels) if rels in names else None
                    )
            for item in zin.infolist():
                data = replacements.pop(item.filename, None) or zin.read(item.filename)
                zout.writestr(item, data)
            for name, data in replacements.items():
                zout.writestr(name, data)
        out_path = Path(out_path)
        out_path.write_bytes(out.getvalue())
        return out_path

    @staticmethod
    def _with_comments_type(data: bytes) -> bytes:
        root = etree.fromstring(data)
        part = "/word/comments.xml"
        for override in root.findall(f"{{{CONTENT_TYPES_NS}}}Override"):
            if override.get("PartName") == part:
                return data
        override = etree.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override")
        override.set("PartName", part)
        override.set("ContentType", COMMENTS_CONTENT_TYPE)
        return _xml(root)

    @staticmethod
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
        return _xml(root)
