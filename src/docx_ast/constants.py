W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PACKAGE_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)

_NS = {"w": W, "r": R, "xml": XML}


def q(name: str) -> str:
    prefix, local = name.split(":")
    return f"{{{_NS[prefix]}}}{local}"


# Heading level → font size (half-points: 32 == 16pt)
_HEADING_SZ = {1: 32, 2: 28, 3: 26, 4: 24, 5: 22, 6: 22}


_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>"""

_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

_DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>"""

# Relationship id (in _DOCUMENT_RELS) that the body's <w:headerReference> points to.
HEADER_RID = "rId3"


def _build_styles() -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<w:styles xmlns:w="{W}">',
        "<w:docDefaults><w:rPrDefault><w:rPr>"
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>',
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>',
    ]
    for lvl in range(1, 7):
        parts.append(
            f'<w:style w:type="paragraph" w:styleId="Heading{lvl}">'
            f'<w:name w:val="heading {lvl}"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr><w:keepNext/><w:outlineLvl w:val="{lvl - 1}"/></w:pPr>'
            f'<w:rPr><w:b/><w:sz w:val="{_HEADING_SZ[lvl]}"/></w:rPr></w:style>'
        )
    parts.append(
        '<w:style w:type="paragraph" w:styleId="FootnoteText">'
        '<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>'
        '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="FootnoteReference">'
        '<w:name w:val="footnote reference"/>'
        '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
        # OrigPage: character style that marks the original page number. It is
        # both the inline marker's look (small, grey) and the anchor the header's
        # STYLEREF field references to show the running "Originalseite N".
        '<w:style w:type="character" w:styleId="OrigPage">'
        '<w:name w:val="OrigPage"/>'
        "</w:style>"
    )
    parts.append("</w:styles>")
    return "".join(parts)


_STYLES = _build_styles()


# Page header: literal "Originalseite " followed by a STYLEREF field that pulls
# the nearest OrigPage-styled text (the original page number) onto each Word
# page, so the header tracks the original pagination through Word's reflow.
_HEADER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{w}">
  <w:p>
    <w:pPr><w:jc w:val="right"/></w:pPr>
    <w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Originalseite </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> STYLEREF "OrigPage" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr><w:t>—</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>""".format(w=W)
