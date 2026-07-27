import tempfile
from pathlib import Path

from docx_ast import (
    ASTDocxSerializer,
    DocxASTExtractor,
    Document,
    Paragraph,
    Text,
    walk,
)


def test_roundtrip():
    doc = Document()
    doc.children.append(Paragraph(children=[Text("Hello world")]))
    doc.children.append(Paragraph(children=[Text("Zweiter Absatz mit § 2 UrhG")]))

    with tempfile.TemporaryDirectory() as d:
        out = Path(d) / "sample.docx"
        ASTDocxSerializer().to_file(doc, out)
        assert out.exists() and out.stat().st_size > 0

        reloaded = DocxASTExtractor().load(out)
        texts = [n.get_text() for n in walk(reloaded) if isinstance(n, Paragraph)]

    assert "Hello world" in texts
    assert any("§ 2 UrhG" in t for t in texts)


if __name__ == "__main__":
    test_roundtrip()
    print("roundtrip OK")
