"""
Rewrite a paragraph's text as Word tracked changes.

The diff (google diff-match-patch) is turned into a list of :class:`Op` whose
offsets all live in the *original* text's coordinate space. Each op then goes to
:mod:`xml_utils.revisions`, which cuts the runs it needs, does the surgery and
hands back a fresh index; applying the ops right to left is what keeps the
offsets of the ones still to come valid. The paragraph's AST is then re-derived
from the edited XML.

Editing therefore assumes a ``w:p`` behind the paragraph — the XML is the source
of truth. A paragraph built from AST nodes alone has nothing to edit and is
rejected; serializing such a document first and reloading it would give it one.
"""

from __future__ import annotations

from dataclasses import dataclass

from diff_match_patch import diff_match_patch

from docx_ast.nodes import Node, Paragraph, Text, walk
from xml_utils.revisions import Rev, delete_range_xml, insert_text_xml
from xml_utils.xml_edit import ParagraphIndex


@dataclass(frozen=True)
class Op:
    """One diff action, in the original text's offset space."""

    kind: str  # "del" | "ins"
    start: int
    end: int  # == start for "ins"
    text: str = ""  # payload, for "ins"


def plan_ops(old: str, new: str) -> list[Op]:
    """Diff ``old`` against ``new`` as offsets into ``old``."""
    dmp = diff_match_patch()
    diffs = dmp.diff_main(old, new)
    dmp.diff_cleanupSemantic(diffs)

    ops: list[Op] = []
    cursor = 0
    for op, chunk in diffs:
        if not chunk:
            continue
        if op == dmp.DIFF_EQUAL:
            cursor += len(chunk)
        elif op == dmp.DIFF_DELETE:
            ops.append(Op("del", cursor, cursor + len(chunk)))
            cursor += len(chunk)
        else:
            ops.append(Op("ins", cursor, cursor, chunk))
    return ops


def apply_ops(
    p_el, ops: list[Op], doc, *, author: str = "", date: str | None = None
) -> None:
    """Apply ``ops`` to a ``w:p`` in place, taking ``w:id``s from ``doc``."""
    if not ops:
        return  # an edit that changes nothing must touch nothing

    # Ids are handed out in reading order even though the ops are applied in
    # reverse, so a document's revision ids ascend the way Word writes them.
    revs = [Rev(doc.next_rev_id(), author, date) for _ in ops]

    # Right to left: an edit at offset q only ever moves text at or after q, so
    # every op still to come keeps the coordinates plan_ops gave it. Where a
    # deletion and an insertion share an offset the deletion goes first — its
    # text stops counting, and the insertion then anchors after the w:del, the
    # order Word writes a substitution in.
    schedule = sorted(
        zip(ops, revs), key=lambda pair: (-pair[0].start, pair[0].kind != "del")
    )

    p_idx = ParagraphIndex(p_el)
    for op, rev in schedule:
        if op.kind == "del":
            p_idx = delete_range_xml(p_idx, op.start, op.end, rev)
        else:
            p_idx = insert_text_xml(p_idx, op.start, op.text, rev)


def alter_paragraph_tracked(
    para: Paragraph,
    new_text: str,
    *,
    author: str = "",
    date: str | None = None,
) -> list[Node]:
    """Rewrite ``para``'s text as tracked changes and return its new children."""
    p_el = para._xml
    ops = plan_ops(ParagraphIndex(p_el).text, new_text)
    if not ops:
        return para.children
    apply_ops(p_el, ops, para.root, author=author, date=date)
    para.refresh()
    return para.children


def alter_text_tracked(
    node: Text,
    new_text: str,
    *,
    author: str = "",
    date: str | None = None,
) -> list[Node]:
    """Rewrite a single ``Text`` node as tracked changes.

    Returns the *paragraph's* children: ``node`` itself does not survive, because
    the paragraph's AST is re-derived from the edited XML.
    """
    para = _enclosing_paragraph(node)
    if para is None:
        raise ValueError(
            "node has no enclosing Paragraph; call link_parents(root) first"
        )
    p_el = para._xml

    base = _char_offset_of(para, node)
    if base is None:
        raise ValueError("node is not reachable from its paragraph (stale parent?)")
    ops = [
        Op(o.kind, o.start + base, o.end + base, o.text)
        for o in plan_ops(node.text, new_text)
    ]
    if not ops:
        return para.children
    apply_ops(p_el, ops, para.root, author=author, date=date)
    para.refresh()
    return para.children


def _enclosing_paragraph(node: Node) -> Paragraph | None:
    while node is not None and not isinstance(node, Paragraph):
        node = node.parent
    return node


def _char_offset_of(para: Paragraph, node: Text) -> int | None:
    pos = 0
    for n in walk(para):
        if n is node:
            return pos
        if isinstance(n, Text):
            pos += len(n.text)
    return None
