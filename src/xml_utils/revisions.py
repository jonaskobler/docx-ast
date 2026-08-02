from __future__ import annotations

import copy
from dataclasses import dataclass

from lxml import etree
from lxml.etree import _Element as E

from xml_utils.element_helpers import (
    hoist,
    nearest_common_ancestor,
    new_text_run,
    preserve_space,
    wrap_siblings,
)
from xml_utils.exceptions import StructureCorruptedError
from xml_utils.ns import DELETED, INSERTED, ln, q
from xml_utils.xml_edit import (
    ParagraphIndex,
    TextInterval,
    apply_cuts,
    split_container,
)

_REVISION_TAGS = (
    "w:ins",
    "w:del",
    "w:moveFrom",
    "w:moveTo",
    "w:moveFromRangeStart",
    "w:moveToRangeStart",
    "w:pPrChange",
    "w:rPrChange",
)


@dataclass(frozen=True)
class Rev:
    """The ``w:id`` / ``w:author`` / ``w:date`` triple of one logical change."""

    id: int
    author: str = ""
    date: str | None = None


def set_rev_attrs(el: E, rev: Rev) -> None:
    el.set(q("w:id"), str(rev.id))
    el.set(q("w:author"), rev.author)
    if rev.date is not None:
        el.set(q("w:date"), rev.date)


def scan_max_rev_id(root: E) -> int:
    """Highest ``w:id`` on any revision element in the tree."""
    tags = {q(t) for t in _REVISION_TAGS}
    top = 0
    for el in root.iter():
        if el.tag in tags:
            try:
                top = max(top, int(el.get(q("w:id"), "0")))
            except ValueError:
                pass  # NOTE: Word only ever writes integers here
    return top


def delete_range_xml(
    p_idx: ParagraphIndex, start: int, end: int, rev: Rev
) -> ParagraphIndex:
    """Mark ``[start, end)`` deleted, one ``w:del`` per container it spans."""
    if end <= start:
        raise ValueError(f"Given span {start},{end} is invalid as end <= start")
    p_idx_cuts_applied = apply_cuts(p_idx, (start, end))
    for runs in _runs_by_container(p_idx_cuts_applied.intervals_in(start, end)):
        # NOTE: We group by container to mark them as deleted
        container = runs[0].getparent()
        if container is None:
            raise StructureCorruptedError()
        if ln(container) in DELETED:
            continue  # NOTE: already gone; deleting it twice means nothing
        _mark_deleted(runs, rev)
    # The deleted text stops being visible, so every offset after `start` moves.
    return ParagraphIndex(p_idx_cuts_applied.p)


def _mark_deleted(runs: list[E], rev: Rev) -> E:
    wrapper = etree.Element(q("w:del"))
    set_rev_attrs(wrapper, rev)
    wrap_siblings(runs[0], runs[-1], wrapper)
    for run in wrapper.iter(q("w:r")):
        for t in run.findall(q("w:t")):
            t.tag = q("w:delText")
            preserve_space(t)
        for t in run.findall(q("w:instrText")):
            t.tag = q("w:delInstrText")
    return wrapper


def _runs_by_container(intervals: list[TextInterval]) -> list[list[E]]:
    """The runs behind ``intervals``, grouped by their container, in order."""
    groups: list[list[E]] = []
    for interval in intervals:
        if groups and groups[-1][-1].getparent() is interval.run_container:
            if groups[-1][-1] is not interval.run:
                groups[-1].append(interval.run)
        else:
            groups.append([interval.run])
    return groups


def insert_text_xml(
    p_idx: ParagraphIndex, at: int, text: str, rev: Rev
) -> ParagraphIndex:
    """Insert ``text`` at offset ``at`` as a ``w:ins``.

    The new text joins the word the caret touches — the neighbour with no space
    between it and the insertion point — and takes both its formatting and its
    container, so typing against a hyperlink extends the link while typing after
    a space does not. When both neighbours are words, or neither is, the left one
    takes it, the way Word continues what precedes the caret.
    """
    if not 0 <= at <= len(p_idx):
        raise ValueError(f"Offset {at} is outside this paragraph")
    p_idx_cuts_applied = apply_cuts(p_idx, (at,))
    p_el = p_idx_cuts_applied.p
    left = p_idx_cuts_applied.interval_ending_at(at)
    right = p_idx_cuts_applied.interval_starting_at(at)

    if _word_ends_at_caret(left):
        # Join the word on the left: its container, just after its run.
        host, parent, before = left, left.run_container, left.run.getnext()
    elif _word_starts_at_caret(right):
        # Join the word on the right: its container, just before its run.
        host, parent, before = right, right.run_container, right.run
    else:
        # A space on both sides: no word to join, so no container is inherited.
        host = left if left is not None else right
        parent, before = _placement(p_el, left, right)

    if parent is None:
        raise StructureCorruptedError()
    parent, before = _out_of_deletions(parent, before)
    run = new_text_run(text, _rpr_for_insert(host, p_el))
    _insert(parent, before, run, rev)
    return ParagraphIndex(p_el)


def _word_ends_at_caret(left: TextInterval | None) -> bool:
    """Is there a word right before the caret — no space"""
    text = (left.t.text or "") if left is not None else ""
    return bool(text) and not text[-1].isspace()


def _word_starts_at_caret(right: TextInterval | None) -> bool:
    """Is there a word right after the caret — no space"""
    text = (right.t.text or "") if right is not None else ""
    return bool(text) and not text[0].isspace()


def _placement(
    p_el: E, left: TextInterval | None, right: TextInterval | None
) -> tuple[E, E | None]:
    """``(parent, before_child)`` when a space separates the caret from both
    neighbours, so there is no word to join.

    The text belongs to no container in particular, so it goes wherever both
    neighbours are visible: their shared container if they have one, the
    paragraph otherwise. That keeps text typed against a leading or trailing
    space outside a hyperlink, while text between two spaces *inside* one stays
    in it.
    """
    if left is not None and right is not None:
        lp, rp = left.run_container, right.run_container
        if lp is None or rp is None:
            raise StructureCorruptedError()
        nca = lp if lp is rp else nearest_common_ancestor(lp, rp)
        parent = nca if nca is not None else p_el
    else:
        parent = p_el  # a paragraph edge

    if right is not None:
        return parent, hoist(right.run, parent)
    if left is not None:
        return parent, hoist(left.run, parent).getnext()
    return parent, None  # empty paragraph


def _out_of_deletions(parent: E, before: E | None) -> tuple[E, E | None]:
    """Live text never goes inside a ``w:del``: step out to just after it."""
    while ln(parent) in DELETED:
        before, parent = parent.getnext(), parent.getparent()
        if parent is None:
            raise StructureCorruptedError()
    return parent, before


def _insert(parent: E, before: E | None, run: E, rev: Rev) -> None:
    """Place ``run`` in ``parent`` before ``before`` (or last), wrapped in ``w:ins``."""
    if ln(parent) in INSERTED and parent.get(q("w:author")) == rev.author:
        # Our own unaccepted insertion: extend it rather than nesting a w:ins.
        _place(parent, before, run)
        return

    ins = etree.Element(q("w:ins"))
    set_rev_attrs(ins, rev)
    ins.append(run)

    if ln(parent) not in INSERTED:
        _place(parent, before, ins)
    elif before is None:  # trailing edge of the other author's insertion
        parent.addnext(ins)
    elif before is parent[0]:  # leading edge
        parent.addprevious(ins)
    else:
        # Strictly inside someone else's insertion: split it so ours sits between
        # the halves rather than nesting w:ins in w:ins. Both halves keep the
        # original id and author — they are still one logical insertion by that
        # author, the same convention as a deletion split across containers.
        left_half, _right_half = split_container(parent, before)
        left_half.addnext(ins)


def _place(parent: E, before: E | None, el: E) -> None:
    if before is not None:
        before.addprevious(el)
    else:
        parent.append(el)


def _rpr_for_insert(host: TextInterval | None, p_el: E) -> E | None:
    """A copy of the ``w:rPr`` the new text should carry.

    The host word's run properties verbatim — including *none*, which is itself
    formatting and must not be papered over with the other neighbour's. With no
    neighbour at all the paragraph mark's run properties stand in, which is what
    a user typing into an empty paragraph would get.
    """
    if host is not None:
        rPr = host.run.find(q("w:rPr"))
    else:
        pPr = p_el.find(q("w:pPr"))
        rPr = pPr.find(q("w:rPr")) if pPr is not None else None
    return copy.deepcopy(rPr) if rPr is not None else None
