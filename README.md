# docx-ast

A lossless Word (`.docx` / OOXML) ↔ AST engine for Python. It reads a Word
document into a small, typed abstract syntax tree, lets you edit / highlight /
comment / apply tracked-change edits on that tree, and writes it back out while
preserving the surrounding document.

It has no dependencies beyond `lxml` and `diff-match-patch` — it manipulates the
OOXML XML directly rather than going through `python-docx`.

## Install

From git (no PyPI release):

```bash
uv pip install "git+https://github.com/jonaskobler/docx-ast"
# or
pip install "git+https://github.com/jonaskobler/docx-ast"
```

## Quick start

```python
from docx_ast import DocxASTExtractor, Paragraph, walk

extractor = DocxASTExtractor()
doc = extractor.load("input.docx")

for para in (n for n in walk(doc) if isinstance(n, Paragraph)):
    print(para.get_text())

# highlight a span and write the document back out
some_paragraph.highlight("BeckOK UrhR Rn. 53", color="cyan")
extractor.save("output.docx")
```

Build a document from scratch:

```python
from docx_ast import ASTDocxSerializer, Document, Paragraph, Text

doc = Document()
doc.children.append(Paragraph(children=[Text("Hello world")]))
ASTDocxSerializer().to_file(doc, "hello.docx")
```

## Public API

- `DocxASTExtractor` — `.load(path)` → `Document`, `.save(out_path)`.
- `ASTDocxSerializer` — `.to_bytes(doc)`, `.to_file(doc, path)` (build a fresh `.docx`).
- `ASTStringSerializer` — render the AST to a readable string.
- Nodes: `Document`, `Paragraph`, `Text`, `Comment`, `Footnote`, `Node`, plus
  `walk`, `link_parents`, `normalize_ws`.
- Tracked changes: `alter_text_tracked`, `alter_paragraph_tracked` (also exposed as
  `.alter_tracked(...)` on `Text` / `Paragraph`).

## Tests

```bash
uv run --extra test pytest
```

## License

MIT
