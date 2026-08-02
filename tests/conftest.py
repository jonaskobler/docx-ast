"""Helpers for building the XML the docx_ast tests read.

Paragraphs are written as XML strings rather than built from nodes: the AST is a
view of the XML, so a test that starts from XML starts where the library does.
"""

import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

from xml_utils.ns import R, W, q

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)

DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)


def run(text: str, *, highlight: str | None = None, bold: bool = False) -> str:
    props = "<w:b/>" if bold else ""
    props += f'<w:highlight w:val="{highlight}"/>' if highlight else ""
    rPr = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'{"<w:r>"}{rPr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def deleted_run(text: str) -> str:
    return f'<w:r><w:delText xml:space="preserve">{text}</w:delText></w:r>'


def para(*pieces: str, pStyle: str | None = None) -> str:
    pPr = f'<w:pPr><w:pStyle w:val="{pStyle}"/></w:pPr>' if pStyle else ""
    return f"<w:p>{pPr}{''.join(pieces)}</w:p>"


def document_xml(*paragraphs: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
        f"{''.join(paragraphs)}<w:sectPr/></w:body></w:document>"
    )


def paragraph_element(*pieces: str, pStyle: str | None = None):
    """A ``w:p`` inside a document, so the ``w:`` / ``r:`` prefixes resolve."""
    root = etree.fromstring(document_xml(para(*pieces, pStyle=pStyle)).encode())
    return root.find(f"{q('w:body')}/{q('w:p')}")


def footnotes_xml(*footnotes: str) -> str:
    """``footnotes`` are ``w:footnote`` fragments; the separators are added here."""
    separators = "".join(
        f'<w:footnote w:type="{kind}" w:id="{fid}"><w:p><w:r><w:{kind}/></w:r></w:p></w:footnote>'
        for fid, kind in ((-1, "separator"), (0, "continuationSeparator"))
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{W}" xmlns:r="{R}">'
        f"{separators}{''.join(footnotes)}</w:footnotes>"
    )


def comments_xml(*comments: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{W}" xmlns:r="{R}">{"".join(comments)}</w:comments>'
    )


def write_docx(
    tmp_path: Path,
    body: str,
    *,
    footnotes: str | None = None,
    comments: str | None = None,
    name: str = "in.docx",
) -> Path:
    """Zip a minimal package around the parts given, and return its path."""
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOCUMENT_RELS,
        "word/document.xml": body,
    }
    if footnotes is not None:
        parts["word/footnotes.xml"] = footnotes
    if comments is not None:
        parts["word/comments.xml"] = comments
        parts["[Content_Types].xml"] = CONTENT_TYPES.replace(
            "</Types>",
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
        )

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name_, data in parts.items():
            z.writestr(name_, data)
    path = tmp_path / name
    path.write_bytes(buf.getvalue())
    return path


def part(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as z:
        return z.read(name)
