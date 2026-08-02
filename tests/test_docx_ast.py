from conftest import (
    comments_xml,
    deleted_run,
    document_xml,
    footnotes_xml,
    para,
    paragraph_element,
    part,
    run,
    write_docx,
)
from lxml import etree

from docx_ast import (
    Del,
    DelText,
    Document,
    Hyperlink,
    Ins,
    Paragraph,
    Text,
    link_parents,
    parse_paragraph,
    walk,
)
from xml_utils.ns import q
from xml_utils.xml_edit import ParagraphIndex


def texts(node) -> list[str]:
    return [n.text for n in walk(node) if isinstance(n, Text)]


def attached(p_el) -> Paragraph:
    """A paragraph under a bare Document, so it can hand out revision ids."""
    doc = Document()
    doc.children.append(parse_paragraph(p_el))
    link_parents(doc)
    return doc.children[0]


# ── parsing ────────────────────────────────────────────────────────────────────


def test_adjacent_runs_merge_into_one_text():
    p = parse_paragraph(paragraph_element(run("Diese "), run("Norm"), run(" gilt")))
    assert texts(p) == ["Diese Norm gilt"]


def test_runs_with_different_highlights_stay_separate():
    p = parse_paragraph(
        paragraph_element(run("Diese "), run("Norm", highlight="cyan"), run(" gilt"))
    )
    assert texts(p) == ["Diese ", "Norm", " gilt"]
    assert [n.highlight for n in p.children] == [None, "cyan", None]


def test_bold_alone_does_not_split_a_text():
    """Formatting the AST does not model must not fragment the view."""
    p = parse_paragraph(paragraph_element(run("Diese "), run("Norm", bold=True)))
    assert texts(p) == ["Diese Norm"]


def test_revisions_nest_under_ins_and_del():
    p = parse_paragraph(
        paragraph_element(
            run("Die "),
            '<w:ins w:id="4" w:author="Me" w:date="2026-01-01T00:00:00Z">'
            + run("neue ")
            + "</w:ins>",
            f'<w:del w:id="5" w:author="Du">{deleted_run("alte ")}</w:del>',
            run("Norm"),
        )
    )
    kinds = [type(n).__name__ for n in p.children]
    assert kinds == ["Text", "Ins", "Del", "Text"]

    ins, dele = p.children[1], p.children[3 - 1]
    assert isinstance(ins, Ins) and (ins.id, ins.author) == (4, "Me")
    assert ins.date == "2026-01-01T00:00:00Z"
    assert isinstance(dele, Del) and (dele.id, dele.author, dele.date) == (
        5,
        "Du",
        None,
    )
    assert texts(ins) == ["neue "]
    assert [n.text for n in walk(dele) if isinstance(n, DelText)] == ["alte "]


def test_hyperlink_keeps_its_children_and_attributes():
    p = parse_paragraph(
        paragraph_element(
            run("siehe "), f'<w:hyperlink r:id="rId9">{run("hier")}</w:hyperlink>'
        )
    )
    link = p.children[1]
    assert isinstance(link, Hyperlink)
    assert texts(link) == ["hier"]
    assert link.attrs[q("r:id")] == "rId9"


def test_sdt_and_move_revisions_are_flattened():
    p = parse_paragraph(
        paragraph_element(
            f"<w:sdt><w:sdtContent>{run('aus ')}</w:sdtContent></w:sdt>",
            f'<w:moveTo w:id="7" w:author="Me">{run("dem ")}</w:moveTo>',
            run("Vertrag"),
        )
    )
    # Flattened *and* merged: the reader sees one uninterrupted stretch.
    assert texts(p) == ["aus dem Vertrag"]


def test_footnote_reference_and_comment_markers_keep_their_place():
    p = parse_paragraph(
        paragraph_element(
            '<w:commentRangeStart w:id="3"/>',
            run("Sosnitza"),
            '<w:commentRangeEnd w:id="3"/>',
            '<w:r><w:footnoteReference w:id="1"/></w:r>',
        )
    )
    assert [type(n).__name__ for n in p.children] == [
        "CommentRangeStart",
        "Text",
        "CommentRangeEnd",
        "FootnoteReference",
    ]
    assert p.children[0].id == 3 and p.children[3].id == 1


def test_paragraph_style_is_read():
    p = parse_paragraph(paragraph_element(run("Titel"), pStyle="Heading1"))
    assert p.pStyle == "Heading1"
    assert parse_paragraph(paragraph_element(run("x"))).pStyle is None


# ── the one offset space ───────────────────────────────────────────────────────


def test_get_text_matches_the_index_edits_are_resolved_against():
    p_el = paragraph_element(
        run("Die "),
        f'<w:ins w:id="1" w:author="Me">{run("neue ")}</w:ins>',
        f'<w:del w:id="2" w:author="Du">{deleted_run("alte ")}</w:del>',
        f'<w:hyperlink r:id="rId9">{run("Norm")}</w:hyperlink>',
        f"<w:sdt><w:sdtContent>{run(' gilt')}</w:sdtContent></w:sdt>",
    )
    para_ = parse_paragraph(p_el)
    assert para_.get_text() == ParagraphIndex(p_el).text == "Die neue Norm gilt"


def test_deleted_text_is_not_part_of_the_text():
    p_el = paragraph_element(
        run("Die "),
        f'<w:del w:id="2" w:author="Du">{deleted_run("alte ")}</w:del>',
        run("Norm"),
    )
    assert parse_paragraph(p_el).get_text() == "Die Norm"


# ── highlighting ───────────────────────────────────────────────────────────────


def test_highlight_marks_only_the_matched_span():
    p = parse_paragraph(paragraph_element(run("Diese Norm gilt")))
    p.highlight("Norm", color="cyan")
    assert [(n.text, n.highlight) for n in p.children] == [
        ("Diese ", None),
        ("Norm", "cyan"),
        (" gilt", None),
    ]


def test_highlight_marks_every_occurrence_across_runs():
    p = parse_paragraph(paragraph_element(run("Norm und "), run("Norm")))
    p.highlight("Norm")
    assert [(n.text, n.highlight) for n in p.children] == [
        ("Norm", "yellow"),
        (" und ", None),
        ("Norm", "yellow"),
    ]


def test_highlight_spans_a_formatting_boundary():
    """The substring crosses two runs, so it must be cut, not missed.

    Both halves end up highlighted the same, so the view shows one span again —
    the run boundary the italics created is not one the AST reports.
    """
    p = parse_paragraph(
        paragraph_element(run("Beck"), run("OK", bold=True), run(" UrhR"))
    )
    p.highlight("BeckOK")
    assert [(n.text, n.highlight) for n in p.children] == [
        ("BeckOK", "yellow"),
        (" UrhR", None),
    ]
    assert p.get_text() == "BeckOK UrhR"


def test_highlighting_a_missing_substring_changes_nothing():
    p_el = paragraph_element(run("Diese Norm"))
    before = etree.tostring(p_el)
    parse_paragraph(p_el).highlight("fehlt")
    assert etree.tostring(p_el) == before


# ── tracked changes ────────────────────────────────────────────────────────────


def test_alter_tracked_rewrites_the_paragraph_as_a_revision():
    p = attached(paragraph_element(run("Die alte Norm")))
    p.alter_tracked("Die neue Norm", author="Me")

    # The exact split is diff-match-patch's business; what matters is that the
    # new text reads back, the old text is kept as a deletion, and both halves
    # are attributed.
    assert p.get_text() == "Die neue Norm"
    ins = next(n for n in walk(p) if isinstance(n, Ins))
    dele = next(n for n in walk(p) if isinstance(n, Del))
    assert ins.author == dele.author == "Me"
    assert "".join(texts(p)) == "Die neue Norm"
    assert "".join(n.text for n in walk(dele) if isinstance(n, DelText)) == "alt"


def test_alter_tracked_on_a_single_text_node():
    """Editing one node offsets the diff by where that node starts."""
    p = attached(
        paragraph_element(
            run("Die "),
            f'<w:hyperlink r:id="rId9">{run("alte")}</w:hyperlink>',
            run(" Norm"),
        )
    )
    node = next(n for n in walk(p) if isinstance(n, Text) and n.text == "alte")
    node.alter_tracked("neue", author="Me")
    assert p.get_text() == "Die neue Norm"


def test_revision_ids_do_not_repeat():
    doc = Document()
    doc.children.append(parse_paragraph(paragraph_element(run("eins zwei"))))
    link_parents(doc)
    doc.children[0].alter_tracked("eins drei", author="Me")

    ids = [
        el.get(q("w:id"))
        for el in doc.children[0]._xml.iter()
        if el.tag in (q("w:ins"), q("w:del"))
    ]
    assert len(ids) == len(set(ids))


# ── the package ────────────────────────────────────────────────────────────────


def test_load_reads_body_footnotes_and_comments(tmp_path):
    path = write_docx(
        tmp_path,
        document_xml(
            para(run("Haupttext"), '<w:r><w:footnoteReference w:id="1"/></w:r>')
        ),
        footnotes=footnotes_xml(
            f'<w:footnote w:id="1"><w:p>{run("Sosnitza, Rn. 5")}</w:p></w:footnote>'
        ),
        comments=comments_xml(
            f'<w:comment w:id="0" w:author="Du" w:initials="D"><w:p>{run("Bitte prüfen")}</w:p></w:comment>'
        ),
    )
    doc = Document.load(path)

    assert [p.get_text() for p in doc.children] == ["Haupttext"]
    assert doc.footnotes[1].children[0].get_text() == "Sosnitza, Rn. 5"
    assert doc.comments[0].author == "Du"
    assert doc.comments[0].children[0].get_text() == "Bitte prüfen"


def test_edits_survive_a_save_and_reload(tmp_path):
    path = write_docx(tmp_path, document_xml(para(run("Die alte Norm gilt hier"))))
    doc = Document.load(path)
    doc.children[0].highlight("Norm", color="cyan")
    doc.children[0].add_comment("gilt", "Stimmt das?", author="Me", initials="M")
    doc.children[0].alter_tracked("Die neue Norm gilt hier", author="Me")

    out = doc.save(tmp_path / "out.docx")
    reloaded = Document.load(out)
    p = reloaded.children[0]

    assert p.get_text() == "Die neue Norm gilt hier"
    assert any(n.highlight == "cyan" for n in walk(p) if isinstance(n, Text))
    assert any(isinstance(n, Ins) for n in walk(p))
    assert reloaded.comments[0].children[0].get_text() == "Stimmt das?"


def test_untouched_paragraphs_are_left_byte_identical(tmp_path):
    path = write_docx(
        tmp_path,
        document_xml(para(run("Erster Absatz")), para(run("Zweiter Absatz"))),
    )
    doc = Document.load(path)
    untouched_before = etree.tostring(doc.children[1]._xml)
    doc.children[0].highlight("Erster")

    reloaded = Document.load(doc.save(tmp_path / "out.docx"))
    assert etree.tostring(reloaded.children[1]._xml) == untouched_before


def test_an_existing_comment_is_not_rewritten(tmp_path):
    """Comments are appended to the part, never re-serialized from the AST, so
    everything the AST does not model survives."""
    existing = (
        '<w:comment w:id="0" w:author="Du" w:initials="D" w:date="2020-01-01T00:00:00Z">'
        f"<w:p>{run('Erste Zeile', bold=True)}</w:p><w:p>{run('Antwort')}</w:p>"
        "</w:comment>"
    )
    path = write_docx(
        tmp_path,
        document_xml(para(run("Die Norm gilt"))),
        comments=comments_xml(existing),
    )
    doc = Document.load(path)
    before = etree.tostring(doc._cmt_root.find(q("w:comment")))

    doc.children[0].add_comment("Norm", "Neue Anmerkung", author="Me")
    reloaded = Document.load(doc.save(tmp_path / "out.docx"))

    assert etree.tostring(reloaded._cmt_root.find(q("w:comment"))) == before
    assert sorted(reloaded.comments) == [0, 1]
    assert reloaded.comments[1].children[0].get_text() == "Neue Anmerkung"


def test_a_comments_part_is_created_when_the_document_has_none(tmp_path):
    path = write_docx(tmp_path, document_xml(para(run("Die Norm gilt"))))
    doc = Document.load(path)
    assert doc._cmt_root is None

    doc.children[0].add_comment("Norm", "Anmerkung", author="Me")
    out = doc.save(tmp_path / "out.docx")

    assert b"/word/comments.xml" in part(out, "[Content_Types].xml")
    assert b"comments.xml" in part(out, "word/_rels/document.xml.rels")
    assert Document.load(out).comments[0].children[0].get_text() == "Anmerkung"


def test_comment_ids_start_after_the_ones_already_used(tmp_path):
    path = write_docx(
        tmp_path,
        document_xml(para('<w:commentRangeStart w:id="7"/>', run("Die Norm gilt"))),
        comments=comments_xml(
            f'<w:comment w:id="7" w:author="Du"><w:p>{run("alt")}</w:p></w:comment>'
        ),
    )
    doc = Document.load(path)
    assert doc.children[0].add_comment("Norm", "neu") == 8


def test_revision_ids_start_after_the_ones_in_the_footnotes(tmp_path):
    """The seed scans every part: a footnote can already carry a revision."""
    path = write_docx(
        tmp_path,
        document_xml(para(run("Haupttext"))),
        footnotes=footnotes_xml(
            '<w:footnote w:id="1"><w:p>'
            f'<w:ins w:id="42" w:author="Du">{run("Sosnitza")}</w:ins>'
            "</w:p></w:footnote>"
        ),
    )
    doc = Document.load(path)
    assert doc.next_rev_id() == 43


def test_comments_and_highlights_work_inside_a_footnote(tmp_path):
    path = write_docx(
        tmp_path,
        document_xml(para(run("Haupttext"))),
        footnotes=footnotes_xml(
            f'<w:footnote w:id="1"><w:p>{run("Sosnitza, Rn. 5")}</w:p></w:footnote>'
        ),
    )
    doc = Document.load(path)
    fn_para = doc.footnotes[1].children[0]
    fn_para.highlight("Rn. 5", color="green")
    assert fn_para.add_comment("Sosnitza", "Welche Auflage?") == 0

    reloaded = Document.load(doc.save(tmp_path / "out.docx"))
    fn = reloaded.footnotes[1].children[0]
    assert any(n.highlight == "green" for n in walk(fn) if isinstance(n, Text))
    assert reloaded.comments[0].children[0].get_text() == "Welche Auflage?"


def test_to_str_renders_the_tree(tmp_path):
    path = write_docx(tmp_path, document_xml(para(run("Die Norm"), pStyle="Heading1")))
    doc = Document.load(path)
    doc.children[0].highlight("Norm", color="cyan")

    rendered = doc.to_str()
    assert "Paragraph(Heading1)" in rendered
    assert "Text('Norm', highlight='cyan')" in rendered
