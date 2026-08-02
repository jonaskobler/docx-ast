from lxml import etree
from testemate.expect_fixture import expect  # noqa: F401  — pytest fixture

from xml_utils.element_helpers import wrap_siblings
from xml_utils.ns import R, W, q
from xml_utils.revisions import Rev, delete_range_xml
from xml_utils.xml_edit import ParagraphIndex, apply_cuts

REV = Rev(1, "Me")


def paragraph(*pieces: str):
    """A ``w:p`` made of ``pieces``, inside a document so ``w:``/``r:`` resolve."""
    xml = (
        f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
        f"<w:p>{''.join(pieces)}</w:p>"
        "</w:body></w:document>"
    )
    return etree.fromstring(xml).find(f"{q('w:body')}/{q('w:p')}")


def run(text: str, *, bold: bool = False) -> str:
    rPr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f"<w:r>{rPr}<w:t>{text}</w:t></w:r>"


def hyperlink(*pieces: str) -> str:
    return f'<w:hyperlink r:id="rId9">{"".join(pieces)}</w:hyperlink>'


def already_deleted(text: str) -> str:
    """Somebody else's unaccepted deletion — delText, the way Word writes it,
    so the index does not see it (see the NOTE on ``iter_intervals``)."""
    return (
        f'<w:del w:id="9" w:author="A"><w:r><w:delText>{text}</w:delText></w:r></w:del>'
    )


PROOF_ERR = '<w:proofErr w:type="spellStart"/>'
BOOKMARK = '<w:bookmarkStart w:id="1" w:name="x"/>'
COMMENT_RANGE_START = '<w:commentRangeStart w:id="3"/>'
XML_COMMENT = "<!-- a comment -->"
PAGE_FIELD = (
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
)


def the_run(el, text: str):
    """The ``w:r`` whose text is ``text``, at any depth.

    Tests name the run they mean instead of indexing into the tree, so a call
    reads without counting children in the XML above it.
    """
    for r in el.iter(q("w:r")):
        if "".join(t.text or "" for t in r) == text:
            return r
    raise AssertionError(f"no run with text {text!r}")


def new_del():
    """An empty ``w:del``, the wrapper the revision code moves runs into."""
    return etree.Element(q("w:del"))


def show(el, label: str | None = None) -> None:
    """Print ``el``: ``w:`` prefixes, no xmlns noise, no trailing blank line."""
    if label:
        print(f"-- {label} --")
    text = etree.tostring(el, pretty_print=True).decode()
    for decl in (f' xmlns:w="{W}"', f' xmlns:r="{R}"'):
        text = text.replace(decl, "")
    print(text.replace("{%s}" % W, "w:").replace("{%s}" % R, "r:").rstrip())


def attempt(fn, *args) -> None:
    """Call ``fn(*args)``, printing the exception instead of raising it."""
    try:
        fn(*args)
    except Exception as e:
        print(f"{type(e).__name__}: {e}")


def test_moves_the_whole_span_and_leaves_the_rest(expect):
    el = paragraph(run("a"), run("b"), run("c"), run("d"))

    wrap_siblings(the_run(el, "b"), the_run(el, "c"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:r>
    <w:t>a</w:t>
  </w:r>
  <w:del>
    <w:r>
      <w:t>b</w:t>
    </w:r>
    <w:r>
      <w:t>c</w:t>
    </w:r>
  </w:del>
  <w:r>
    <w:t>d</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )


def test_single_element_span(expect):
    el = paragraph(run("a"), run("b"))

    wrap_siblings(the_run(el, "a"), the_run(el, "a"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:del>
    <w:r>
      <w:t>a</w:t>
    </w:r>
  </w:del>
  <w:r>
    <w:t>b</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )


def test_non_run_content_between_the_runs_travels_along(expect):
    el = paragraph(run("a"), COMMENT_RANGE_START, run("b"), run("c"))

    wrap_siblings(the_run(el, "a"), the_run(el, "b"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:del>
    <w:r>
      <w:t>a</w:t>
    </w:r>
    <w:commentRangeStart w:id="3"/>
    <w:r>
      <w:t>b</w:t>
    </w:r>
  </w:del>
  <w:r>
    <w:t>c</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )


def test_proof_err_is_dropped(expect):
    el = paragraph(run("a"), PROOF_ERR, run("b"))

    wrap_siblings(the_run(el, "a"), the_run(el, "b"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:del>
    <w:r>
      <w:t>a</w:t>
    </w:r>
    <w:r>
      <w:t>b</w:t>
    </w:r>
  </w:del>
</w:p>
""",
        debug=True,
    )


def test_bookmarks_are_hoisted_in_front_of_the_wrapper(expect):
    el = paragraph(run("a"), BOOKMARK, run("b"))

    wrap_siblings(the_run(el, "a"), the_run(el, "b"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:bookmarkStart w:id="1" w:name="x"/>
  <w:del>
    <w:r>
      <w:t>a</w:t>
    </w:r>
    <w:r>
      <w:t>b</w:t>
    </w:r>
  </w:del>
</w:p>
""",
        debug=True,
    )


def test_comments_stay_outside(expect):
    el = paragraph(run("a"), XML_COMMENT, run("b"))

    wrap_siblings(the_run(el, "a"), the_run(el, "b"), new_del())

    show(el)
    expect(
        """\
<w:p>
  <w:del>
    <w:r>
      <w:t>a</w:t>
    </w:r>
    <w:r>
      <w:t>b</w:t>
    </w:r>
  </w:del>
  <!-- a comment -->
</w:p>
""",
        debug=True,
    )


def test_rejects_a_last_that_is_not_a_sibling(expect):
    """A bad pair must raise before anything moves, never half-wrap."""
    el = paragraph(run("a"), hyperlink(run("b")))

    # "b" lives inside the hyperlink, so it is no sibling of "a"
    attempt(wrap_siblings, the_run(el, "a"), the_run(el, "b"), new_del())

    show(el, "paragraph afterwards")
    expect(
        """\
ValueError: Element is not a child of this node.
-- paragraph afterwards --
<w:p>
  <w:r>
    <w:t>a</w:t>
  </w:r>
  <w:hyperlink r:id="rId9">
    <w:r>
      <w:t>b</w:t>
    </w:r>
  </w:hyperlink>
</w:p>
""",
        debug=True,
    )


def test_rejects_a_span_from_another_tree(expect):
    el = paragraph(run("a"), run("b"))
    stranger = the_run(paragraph(run("z")), "z")

    attempt(wrap_siblings, the_run(el, "a"), stranger, new_del())

    show(el, "paragraph afterwards")
    expect(
        """\
ValueError: Element is not a child of this node.
-- paragraph afterwards --
<w:p>
  <w:r>
    <w:t>a</w:t>
  </w:r>
  <w:r>
    <w:t>b</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )


def test_rejects_a_reversed_span(expect):
    el = paragraph(run("a"), run("b"), run("c"))

    attempt(wrap_siblings, the_run(el, "c"), the_run(el, "a"), new_del())

    show(el, "paragraph afterwards")
    expect(
        """\
ValueError: last comes before first
-- paragraph afterwards --
<w:p>
  <w:r>
    <w:t>a</w:t>
  </w:r>
  <w:r>
    <w:t>b</w:t>
  </w:r>
  <w:r>
    <w:t>c</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )


def test_rejects_wrapping_the_root(expect):
    document = paragraph(run("a")).getroottree().getroot()

    attempt(wrap_siblings, document, document, new_del())

    expect(
        """\
ValueError: Trying to wrap a root
""",
        debug=True,
    )


# --------------------------------------------------------------------------- #
# What the delete path builds on top of it
# --------------------------------------------------------------------------- #


def test_delete_wraps_in_place_and_retags(expect):
    el = paragraph(run("the quick brown fox"))

    idx = delete_range_xml(ParagraphIndex(el), 4, 10, REV)

    show(el)
    print(f"visible text: {idx.text!r}")  # delText is invisible to the index
    expect(
        """\
<w:p>
  <w:r>
    <w:t xml:space="preserve">the </w:t>
  </w:r>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">quick </w:delText>
    </w:r>
  </w:del>
  <w:r>
    <w:t xml:space="preserve">brown fox</w:t>
  </w:r>
</w:p>
visible text: 'the brown fox'
""",
        debug=True,
    )


def test_delete_across_a_hyperlink_is_several_dels_with_one_id(expect):
    el = paragraph(run("abc"), hyperlink(run("def")), run("ghi"))

    idx = delete_range_xml(ParagraphIndex(el), 2, 7, REV)

    show(el)
    print(f"visible text: {idx.text!r}")
    expect(
        """\
<w:p>
  <w:r>
    <w:t xml:space="preserve">ab</w:t>
  </w:r>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">c</w:delText>
    </w:r>
  </w:del>
  <w:hyperlink r:id="rId9">
    <w:del w:id="1" w:author="Me">
      <w:r>
        <w:delText xml:space="preserve">def</w:delText>
      </w:r>
    </w:del>
  </w:hyperlink>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">g</w:delText>
    </w:r>
  </w:del>
  <w:r>
    <w:t xml:space="preserve">hi</w:t>
  </w:r>
</w:p>
visible text: 'abhi'
""",
        debug=True,
    )


def test_delete_beside_an_existing_deletion(expect):
    """The old w:del keeps its id and author; ours wraps only what is visible.

    "gone" is delText, so it holds no offsets at all — [0,4) is "kept".
    """
    el = paragraph(already_deleted("gone"), run("kept"))

    show(el, "before")
    delete_range_xml(ParagraphIndex(el), 0, 4, REV)
    show(el, "after deleting [0,4)")
    expect(
        """\
-- before --
<w:p>
  <w:del w:id="9" w:author="A">
    <w:r>
      <w:delText>gone</w:delText>
    </w:r>
  </w:del>
  <w:r>
    <w:t>kept</w:t>
  </w:r>
</w:p>
-- after deleting [0,4) --
<w:p>
  <w:del w:id="9" w:author="A">
    <w:r>
      <w:delText>gone</w:delText>
    </w:r>
  </w:del>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">kept</w:delText>
    </w:r>
  </w:del>
</w:p>
""",
        debug=True,
    )


def test_delete_retags_field_instructions(expect):
    el = paragraph(run("page "), PAGE_FIELD, run(" end"))

    delete_range_xml(ParagraphIndex(el), 0, 9, REV)

    show(el)
    expect(
        """\
<w:p>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">page </w:delText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>
    </w:r>
    <w:r>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r>
      <w:delText xml:space="preserve"> end</w:delText>
    </w:r>
  </w:del>
</w:p>
""",
        debug=True,
    )


def test_delete_keeps_run_formatting_on_both_sides_of_a_split(expect):
    el = paragraph(run("abc"), run("def", bold=True))

    idx = delete_range_xml(ParagraphIndex(el), 1, 5, REV)

    show(el)
    print(f"visible text: {idx.text!r}")
    expect(
        """\
<w:p>
  <w:r>
    <w:t xml:space="preserve">a</w:t>
  </w:r>
  <w:del w:id="1" w:author="Me">
    <w:r>
      <w:delText xml:space="preserve">bc</w:delText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:b/>
      </w:rPr>
      <w:delText xml:space="preserve">de</w:delText>
    </w:r>
  </w:del>
  <w:r>
    <w:rPr>
      <w:b/>
    </w:rPr>
    <w:t xml:space="preserve">f</w:t>
  </w:r>
</w:p>
visible text: 'af'
""",
        debug=True,
    )


def test_cut_on_an_existing_boundary_adds_nothing(expect):
    el = paragraph(run("abc"), run("def"))

    show(el, "before")
    apply_cuts(ParagraphIndex(el), [0, 3, 6])
    show(el, "after cutting at 0, 3, 6")
    expect(
        """\
-- before --
<w:p>
  <w:r>
    <w:t>abc</w:t>
  </w:r>
  <w:r>
    <w:t>def</w:t>
  </w:r>
</w:p>
-- after cutting at 0, 3, 6 --
<w:p>
  <w:r>
    <w:t>abc</w:t>
  </w:r>
  <w:r>
    <w:t>def</w:t>
  </w:r>
</w:p>
""",
        debug=True,
    )
