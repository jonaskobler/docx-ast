"""
Serialize an AST ``Document`` into a Word ``.docx`` package.

The body (``word/document.xml``) and the footnotes (``word/footnotes.xml``) are
built with ``lxml`` straight from the AST; each node type is handled by a
``functools.singledispatchmethod`` registration. The remaining OPC parts
(content types, relationships, styles) are static boilerplate, and the whole
package is zipped together.

    from docx_ast.ast_to_docx import ASTDocxSerializer
    ASTDocxSerializer().to_file(document, "out.docx")
"""

from __future__ import annotations

import copy
import zipfile
from functools import singledispatchmethod
from io import BytesIO
from pathlib import Path

from lxml import etree

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
    ParagraphStyle,
    RunStyle,
    Sdt,
    Tab,
    Text,
)
from docx_ast.constants import (
    q,
    _CONTENT_TYPES,
    _HEADER,
    _STYLES,
    _DOCUMENT_RELS,
    _ROOT_RELS,
    HEADER_RID,
    W,
    R,
)


class ASTDocxSerializer:
    def __init__(self) -> None:
        self._styles: dict = {}

    def to_bytes(self, document: Document) -> bytes:
        self._styles = document.styles
        document_xml = self._tostring(self._document_xml(document))
        footnotes_xml = self._tostring(self._footnotes_xml(document))
        return self._package(document_xml, footnotes_xml)

    def to_file(self, document: Document, path: str | Path) -> Path:
        path = Path(path)
        path.write_bytes(self.to_bytes(document))
        return path

    def _rPr_for_style(self, style_id: str | None) -> etree._Element | None:
        """Return a <w:rPr> element for the given style_id, or None."""
        if not style_id:
            return None
        style = self._styles.get(style_id)
        if not isinstance(style, RunStyle):
            return None
        if style._raw is not None:
            return copy.deepcopy(style._raw)
        # Build from explicit fields when no raw element is available.
        rPr = etree.Element(q("w:rPr"))
        if style.rStyle:
            etree.SubElement(rPr, q("w:rStyle")).set(q("w:val"), style.rStyle)
        if style.bold:
            etree.SubElement(rPr, q("w:b"))
        if style.italic:
            etree.SubElement(rPr, q("w:i"))
        if style.font_name:
            e = etree.SubElement(rPr, q("w:rFonts"))
            e.set(q("w:ascii"), style.font_name)
            e.set(q("w:hAnsi"), style.font_name)
        if style.font_size:
            etree.SubElement(rPr, q("w:sz")).set(q("w:val"), str(style.font_size))
        if style.color:
            etree.SubElement(rPr, q("w:color")).set(q("w:val"), style.color)
        return rPr if len(rPr) > 0 else None

    @singledispatchmethod
    def _emit(self, node: Node, parent: etree._Element) -> None:
        raise TypeError(f"No serializer registered for {type(node).__name__}")

    @_emit.register(Paragraph)
    def _(self, node: Paragraph, parent: etree._Element) -> None:
        p = etree.SubElement(parent, q("w:p"))
        if node.style_id:
            style = self._styles.get(node.style_id)
            if isinstance(style, ParagraphStyle):
                if style._raw is not None:
                    p.append(copy.deepcopy(style._raw))
                elif style.pStyle:
                    pPr = etree.SubElement(p, q("w:pPr"))
                    etree.SubElement(pPr, q("w:pStyle")).set(q("w:val"), style.pStyle)
        for child in node.children:
            self._emit(child, p)

    @_emit.register(PageBreak)
    def _(self, node: PageBreak, parent: etree._Element) -> None:
        self._page_marker(parent, node.page)

    def _page_marker(self, parent: etree._Element, page: int | None) -> None:
        """Visible transition marker ``|N-1→N|`` at a break, with *two* OrigPage
        anchors in order: the leaving page (``N-1``) then the beginning page
        (``N``). The header's STYLEREF shows the first on-page anchor, else the
        nearest one above:

        * On the break page (top = N-1) the first on-page anchor is N-1 → shows N-1.
        * On a following break-less page the backward search finds the most recent
          anchor, now the beginning page N → shows N (fixes the interior page).

        Anchors must be visible: STYLEREF does not read hidden (vanished) text.
        """
        p = self._centered_paragraph(parent)
        if page is None:
            self._make_run_with_text(p, "|—|")
            return
        self._make_run_with_text(p, "-------------")
        self._anchor_run(p, page - 1)  # leaving page → first-on-page for break page
        self._make_run_with_text(p, "→")
        self._anchor_run(p, page)  # beginning page → backward search for later pages
        self._make_run_with_text(p, "-------------")

    def _anchor_paragraph(self, parent: etree._Element, value: int) -> None:
        """Standalone OrigPage anchor (document start/end), so the first and last
        pages resolve before the first / after the last break."""
        p = self._centered_paragraph(parent)
        self._make_run_with_text(p, "|")
        self._anchor_run(p, value)
        self._make_run_with_text(p, "|")

    @staticmethod
    def _centered_paragraph(parent: etree._Element) -> etree._Element:
        p = etree.SubElement(parent, q("w:p"))
        etree.SubElement(etree.SubElement(p, q("w:pPr")), q("w:jc")).set(
            q("w:val"), "center"
        )
        return p

    @staticmethod
    def _anchor_run(p: etree._Element, value: int) -> None:
        """Visible run (OrigPage character style) the header's STYLEREF reads."""
        r = etree.SubElement(p, q("w:r"))
        etree.SubElement(etree.SubElement(r, q("w:rPr")), q("w:rStyle")).set(
            q("w:val"), "OrigPage"
        )
        t = etree.SubElement(r, q("w:t"))
        t.set(q("xml:space"), "preserve")
        t.text = str(value)

    @staticmethod
    def _make_run_with_text(p: etree._Element, text: str) -> None:
        r = etree.SubElement(p, q("w:r"))
        t = etree.SubElement(r, q("w:t"))
        t.set(q("xml:space"), "preserve")
        t.text = text

    @_emit.register(Text)
    def _(self, node: Text, parent: etree._Element) -> None:
        if not node.text:
            return
        r = etree.SubElement(parent, q("w:r"))
        rPr = self._rPr_for_style(node.style_id)
        if rPr is not None:
            r.append(rPr)
        t = etree.SubElement(r, q("w:t"))
        t.set(q("xml:space"), "preserve")
        t.text = node.text

    @_emit.register(DelText)
    def _(self, node: DelText, parent: etree._Element) -> None:
        if not node.text:
            return
        r = etree.SubElement(parent, q("w:r"))
        rPr = self._rPr_for_style(node.style_id)
        if rPr is not None:
            r.append(rPr)
        t = etree.SubElement(r, q("w:delText"))
        t.set(q("xml:space"), "preserve")
        t.text = node.text

    @_emit.register(Ins)
    def _(self, node: Ins, parent: etree._Element) -> None:
        ins = etree.SubElement(parent, q("w:ins"))
        self._set_rev_attrs(ins, node)
        for child in node.children:
            self._emit(child, ins)

    @_emit.register(Del)
    def _(self, node: Del, parent: etree._Element) -> None:
        # Word stores deleted text in <w:delText>; use DelText leaves inside a Del.
        dele = etree.SubElement(parent, q("w:del"))
        self._set_rev_attrs(dele, node)
        for child in node.children:
            self._emit(child, dele)

    @staticmethod
    def _set_rev_attrs(el: etree._Element, node: Ins | Del) -> None:
        """Apply the shared tracked-change revision attributes. w:id and
        w:author are required by the schema; w:date is optional."""
        el.set(q("w:id"), str(node.id))
        el.set(q("w:author"), node.author)
        if node.date is not None:
            el.set(q("w:date"), node.date)

    @_emit.register(MoveFrom)
    def _(self, node: MoveFrom, parent: etree._Element) -> None:
        el = etree.SubElement(parent, q("w:moveFrom"))
        self._set_move_attrs(el, node)
        for child in node.children:
            self._emit(child, el)

    @_emit.register(MoveTo)
    def _(self, node: MoveTo, parent: etree._Element) -> None:
        el = etree.SubElement(parent, q("w:moveTo"))
        self._set_move_attrs(el, node)
        for child in node.children:
            self._emit(child, el)

    @staticmethod
    def _set_move_attrs(el: etree._Element, node: MoveFrom | MoveTo) -> None:
        if node.attrs:
            for k, v in node.attrs.items():
                el.set(k, v)
            return
        el.set(q("w:id"), str(node.id))
        el.set(q("w:author"), node.author)
        if node.date is not None:
            el.set(q("w:date"), node.date)

    @_emit.register(MoveRangeMarker)
    def _(self, node: MoveRangeMarker, parent: etree._Element) -> None:
        el = etree.SubElement(parent, q(f"w:{node.kind}"))
        for k, v in node.attrs.items():
            el.set(k, v)

    @_emit.register(Hyperlink)
    def _(self, node: Hyperlink, parent: etree._Element) -> None:
        h = etree.SubElement(parent, q("w:hyperlink"))
        for k, v in node.attrs.items():
            h.set(k, v)
        for child in node.children:
            self._emit(child, h)

    @_emit.register(Sdt)
    def _(self, node: Sdt, parent: etree._Element) -> None:
        sdt = etree.SubElement(parent, q("w:sdt"))
        for pr in node.pr:
            sdt.append(copy.deepcopy(pr))
        content = etree.SubElement(sdt, q("w:sdtContent"))
        for child in node.children:
            self._emit(child, content)

    @_emit.register(Tab)
    def _(self, node: Tab, parent: etree._Element) -> None:
        r = etree.SubElement(parent, q("w:r"))
        rPr = self._rPr_for_style(node.style_id)
        if rPr is not None:
            r.append(rPr)
        etree.SubElement(r, q("w:tab"))

    @_emit.register(FieldChar)
    def _(self, node: FieldChar, parent: etree._Element) -> None:
        r = etree.SubElement(parent, q("w:r"))
        rPr = self._rPr_for_style(node.style_id)
        if rPr is not None:
            r.append(rPr)
        etree.SubElement(r, q("w:fldChar")).set(q("w:fldCharType"), node.kind)

    @_emit.register(InstrText)
    def _(self, node: InstrText, parent: etree._Element) -> None:
        r = etree.SubElement(parent, q("w:r"))
        rPr = self._rPr_for_style(node.style_id)
        if rPr is not None:
            r.append(rPr)
        t = etree.SubElement(r, q("w:instrText"))
        t.set(q("xml:space"), "preserve")
        t.text = node.text

    @_emit.register(FootnoteReference)
    def _(self, node: FootnoteReference, parent: etree._Element) -> None:
        r = etree.SubElement(parent, q("w:r"))
        rPr = etree.SubElement(r, q("w:rPr"))
        etree.SubElement(rPr, q("w:rStyle")).set(q("w:val"), "FootnoteReference")
        etree.SubElement(rPr, q("w:vertAlign")).set(q("w:val"), "superscript")
        ref = etree.SubElement(r, q("w:footnoteReference"))
        ref.set(q("w:id"), str(node.id))

    @_emit.register(CommentRangeStart)
    def _(self, node: CommentRangeStart, parent: etree._Element) -> None:
        etree.SubElement(parent, q("w:commentRangeStart")).set(q("w:id"), str(node.id))

    @_emit.register(CommentRangeEnd)
    def _(self, node: CommentRangeEnd, parent: etree._Element) -> None:
        etree.SubElement(parent, q("w:commentRangeEnd")).set(q("w:id"), str(node.id))
        r = etree.SubElement(parent, q("w:r"))
        rPr = etree.SubElement(r, q("w:rPr"))
        etree.SubElement(rPr, q("w:rStyle")).set(q("w:val"), "CommentReference")
        etree.SubElement(r, q("w:commentReference")).set(q("w:id"), str(node.id))

    @_emit.register(Footnote)
    def _(self, node: Footnote, parent: etree._Element) -> None:
        fn = etree.SubElement(parent, q("w:footnote"))
        fn.set(q("w:id"), str(node.id))
        p = etree.SubElement(fn, q("w:p"))
        pPr = etree.SubElement(p, q("w:pPr"))
        etree.SubElement(pPr, q("w:pStyle")).set(q("w:val"), "FootnoteText")

        ref_run = etree.SubElement(p, q("w:r"))
        ref_rPr = etree.SubElement(ref_run, q("w:rPr"))
        etree.SubElement(ref_rPr, q("w:rStyle")).set(q("w:val"), "FootnoteReference")
        etree.SubElement(ref_run, q("w:footnoteRef"))
        spacer = etree.SubElement(p, q("w:r"))
        spacer_t = etree.SubElement(spacer, q("w:t"))
        spacer_t.set(q("xml:space"), "preserve")
        spacer_t.text = " "

        for child in node.children:
            self._emit(child, p)

    def _document_xml(self, document: Document) -> etree._Element:
        root = etree.Element(q("w:document"), nsmap={"w": W, "r": R})
        body = etree.SubElement(root, q("w:body"))

        # Seed anchor for the page *before* the first break, so content above the
        # first break (and the very first page) resolves in the header. No trailing
        # anchor is needed: each break's "beginning" anchor already covers the page
        # below it, including the final page.
        pbs = [c for c in document.children if isinstance(c, PageBreak)]
        if pbs and pbs[0].page and pbs[0].page > 1:
            self._anchor_paragraph(body, pbs[0].page - 1)

        for child in document.children:
            self._emit(child, body)

        sectPr = etree.SubElement(body, q("w:sectPr"))
        hdr = etree.SubElement(sectPr, q("w:headerReference"))
        hdr.set(q("w:type"), "default")
        hdr.set(q("r:id"), HEADER_RID)
        return root

    def _footnotes_xml(self, document: Document) -> etree._Element:
        root = etree.Element(q("w:footnotes"), nsmap={"w": W, "r": R})
        for fid, kind in ((-1, "separator"), (0, "continuationSeparator")):
            fn = etree.SubElement(root, q("w:footnote"))
            fn.set(q("w:type"), kind)
            fn.set(q("w:id"), str(fid))
            p = etree.SubElement(fn, q("w:p"))
            r = etree.SubElement(p, q("w:r"))
            etree.SubElement(r, q(f"w:{kind}"))
        for fid in sorted(document.footnotes):
            self._emit(document.footnotes[fid], root)
        return root

    def comments_xml(self, comments: list[Comment]) -> etree._Element:
        root = etree.Element(q("w:comments"), nsmap={"w": W, "r": R})
        for comment in comments:
            self._emit_comment_part(comment, root)
        return root

    def _emit_comment_part(self, node: Comment, root: etree._Element) -> None:
        c = etree.SubElement(root, q("w:comment"))
        c.set(q("w:id"), str(node.id))
        c.set(q("w:author"), node.author)
        if node.initials:
            c.set(q("w:initials"), node.initials)
        if node.date is not None:
            c.set(q("w:date"), node.date)
        for para in self._comment_paragraphs(node.content):
            self._emit(para, c)

    @staticmethod
    def _comment_paragraphs(content: list[Node]) -> list[Paragraph]:
        paragraphs: list[Paragraph] = []
        loose: list[Node] = []
        for n in content:
            if isinstance(n, Paragraph):
                if loose:
                    paragraphs.append(Paragraph(children=loose))
                    loose = []
                paragraphs.append(n)
            else:
                loose.append(n)
        if loose:
            paragraphs.append(Paragraph(children=loose))
        return paragraphs or [Paragraph()]

    @staticmethod
    def _tostring(root: etree._Element) -> bytes:
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    @staticmethod
    def _package(document_xml: bytes, footnotes_xml: bytes) -> bytes:
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CONTENT_TYPES)
            z.writestr("_rels/.rels", _ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DOCUMENT_RELS)
            z.writestr("word/styles.xml", _STYLES)
            z.writestr("word/header1.xml", _HEADER)
            z.writestr("word/document.xml", document_xml)
            z.writestr("word/footnotes.xml", footnotes_xml)
        return buf.getvalue()
