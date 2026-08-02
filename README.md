# docx-ast

A Word (`.docx` / OOXML) editing engine for Python. It reads a Word document into
a small, typed abstract syntax tree, lets you highlight, comment and apply
tracked-change edits through that tree, and writes the document back out with
everything it did not touch left alone.

It has no dependencies beyond `lxml` and `diff-match-patch` — it manipulates the
OOXML XML directly rather than going through `python-docx`.

## How editing works

The AST is a **view**, not a copy. Every `Paragraph` keeps a live pointer to its
`w:p` element, and edits are applied to that XML in place; the paragraph's AST is
then re-derived from the result. Anything the AST does not model — bookmarks,
drawings, field codes, unusual run content, tracked changes made by other authors
— stays byte-identical through an edit, and the surrounding document is never
rewritten.

The tree is deliberately reduced: it exists so a reader (typically an LLM) can see
the document as something close to markdown, and it models only what that reader
acts on — text, tracked insertions and deletions, highlights, comment ranges,
footnote references. Everything else stays in the XML, untouched and unmodelled.

Concretely, an edit resolves character offsets against an immutable index of the
paragraph's `w:t` elements, materialises the boundaries it needs as run
boundaries, and then does the XML surgery right-to-left so no offset is ever used
after the tree beneath it has moved. Where a `w:ins` / `w:del` goes depends on
what the affected runs sit inside: text in a hyperlink is edited inside the
`w:hyperlink`, but text inserted at a link's outer edge lands outside it (as in
Word), and re-deleting another author's insertion nests a `w:del` inside their
`w:ins`.

## Install

From git (no PyPI release):

```bash
uv pip install "git+https://github.com/jonaskobler/docx-ast"
# or
pip install "git+https://github.com/jonaskobler/docx-ast"
```

## Quick start

```python
from docx_ast import Document

doc = Document.load("input.docx")

for para in doc.children:
    print(para.get_text())

# footnotes are paragraphs too, and can be edited the same way
for fid, footnote in sorted(doc.footnotes.items()):
    print(fid, footnote.children[0].get_text())

para = doc.children[0]
para.highlight("BeckOK UrhR Rn. 53", color="cyan")
para.add_comment("Rn. 53", "Auflage prüfen", author="Me", initials="M")
para.alter_tracked(para.get_text().replace("49", "50"), author="Me")

doc.save("output.docx")
```

## Public API

- `Document` — `Document.load(path)`, `.save(out_path)`. It is both the loaded
  package and the root of the tree: `.children` are the body paragraphs,
  `.footnotes` and `.comments` are keyed by the ids Word uses.
- `Paragraph` — `.get_text()`, `.highlight(substring, color)`,
  `.highlight_spans(spans, color)`, `.add_comment(substring, note, ...)`,
  `.alter_tracked(new_text, ...)`, and `.refresh()` to re-derive the children
  from the `w:p` (every edit above already calls it for you).
- Nodes: `Text`, `DelText`, `Ins`, `Del`, `Hyperlink`, `Tab`,
  `FootnoteReference`, `CommentRangeStart`, `CommentRangeEnd`, `Footnote`,
  `Comment`, `Node`, plus `walk`, `link_parents`, `normalize_ws`.
- Tracked changes: `alter_text_tracked`, `alter_paragraph_tracked` (also exposed as
  `.alter_tracked(...)` on `Text` / `Paragraph`), plus `plan_ops` / `Op` if you
  want the diff as offsets without applying it.
- Parsing / XML: `parse_paragraph`, `to_str` for a readable dump of the tree, and
  `ParagraphIndex` for the character-offset → `w:t` mapping.

Both `alter_*_tracked` functions return the *paragraph's* children. `Text` nodes
do not survive an edit, since the paragraph's AST is re-derived from the XML —
re-fetch the nodes you need afterwards rather than holding them across a call.
`Paragraph` objects themselves stay valid across any number of edits.

Character offsets — the ones `highlight_spans` takes and `plan_ops` returns — are
positions in `Paragraph.get_text()`, which counts visible text only: deleted text
(`w:delText`) and tabs contribute no characters.

## Tests

```bash
uv run --extra test pytest
```

## License

MIT
