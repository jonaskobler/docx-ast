from __future__ import annotations

import copy
from bisect import bisect_right
from collections.abc import Iterable, Iterator
from dataclasses import dataclass

from lxml import etree
from lxml.etree import _Element as E

from xml_utils.element_helpers import (
    comment_reference_run,
    preserve_space,
    range_marker,
    set_rpr_prop,
)
from xml_utils.exceptions import StructureCorruptedError
from xml_utils.ns import RUN_CONTAINERS, is_element, ln, q


@dataclass(frozen=True)
class TextInterval:
    """One ``w:t`` element and the character range it covers. No delText cause we dont care, see below"""

    t: E
    run: E
    start: int
    end: int

    @property
    def run_container(self) -> E | None:
        return self.run.getparent()

    def anchor_before(self, el: E) -> None:
        self.run.addprevious(el)

    def anchor_after(self, el: E) -> None:
        self.run.addnext(el)


def iter_intervals(el: E) -> Iterator[tuple[E, E]]:
    """Every ``(w:t, run)`` pair below ``el``, in document order.

    NOTE: The walk descends into ``w:del`` and ``w:moveFrom`` — they are in
    RUN_CONTAINERS — and skips only the ``delText`` *tag*, so it assumes deleted
    content is tagged ``delText``, the way Word writes it. A ``w:del`` holding a
    ``w:t`` would be indexed as visible text (``docx_to_ast`` reads it as a
    visible Text node too, so the two agree), and revisions.py relies on that not
    happening: no interval's run is inside a deletion, which is why nothing there
    checks for it.
    """
    for child in el:
        if not is_element(child):
            continue
        tag = ln(child)
        if tag == "r":
            for gc in child:
                if is_element(gc) and ln(gc) == "t":
                    # NOTE: We also walk on text, we ignore delText because we dont care about deleted text (for now?)
                    yield gc, child
        elif tag in RUN_CONTAINERS:
            yield from iter_intervals(child)
        elif tag == "sdt":
            content = child.find(q("w:sdtContent"))
            if content is not None:
                yield from iter_intervals(content)


class ParagraphIndex:
    """Immutable snapshot mapping character offsets to ``w:t`` elements. Usage: find w:t elements to split for insertion of tc/etc"""

    __slots__ = ("p", "text_intervals", "text", "_starts")

    def __init__(self, p_el: E) -> None:
        self.p = p_el
        self.text_intervals: list[TextInterval] = []
        pos = 0
        for t_el, run in iter_intervals(p_el):
            length = len(t_el.text or "")  # NOTE: I think w:t's cannot contain no text?
            self.text_intervals.append(TextInterval(t_el, run, pos, pos + length))
            pos += length
        self._starts: list[int] = [s.start for s in self.text_intervals]
        self.text: str = "".join(s.t.text or "" for s in self.text_intervals)

    def __len__(self) -> int:
        return len(self.text)

    def text_interval_at(self, off: int) -> TextInterval | None:
        """The interval containing ``off`` (``start <= off < end``)."""
        i = bisect_right(self._starts, off) - 1
        if i < 0:
            return None
        s = self.text_intervals[i]
        return s if s.end > off else None

    def interval_starting_at(self, off: int) -> TextInterval | None:
        i = bisect_right(self._starts, off) - 1
        if i < 0:
            return None
        s = self.text_intervals[i]
        return s if s.start == off else None

    def interval_ending_at(self, off: int) -> TextInterval | None:
        s = self.text_interval_at(off - 1) if off > 0 else None
        return s if s is not None and s.end == off else None

    def intervals_in(self, a: int, b: int) -> list[TextInterval]:
        """Visible intervals overlapping ``[a, b)``, in document order."""
        return [s for s in self.text_intervals if s.start < b and s.end > a]


def split_wt_at(t_el: E, k: int) -> tuple[E, E]:
    run = t_el.getparent()
    if run is None:
        raise StructureCorruptedError()
    i = run.index(t_el)
    right = copy.deepcopy(
        run
    )  # carries w:rPr and every content child, # TODO: Remove deepcopy?

    text = t_el.text or ""
    t_el.text, right[i].text = text[:k], text[k:]
    preserve_space(t_el)
    preserve_space(right[i])

    for child in list(run)[i + 1 :]:
        run.remove(child)
    for child in list(right)[:i]:
        if ln(child) != "rPr":  # NOTE: Delete everything except the rPr
            right.remove(child)

    run.addnext(right)
    return run, right


def apply_cuts(p_idx: ParagraphIndex, offsets: Iterable[int]) -> ParagraphIndex:
    """Make every offset in ``offsets`` a run boundary, and re-index."""
    for off in sorted(offsets, reverse=True):
        text_interval = p_idx.text_interval_at(off)
        if text_interval is None:
            continue  # NOTE: That should mean that we are out of bounds
        if off == text_interval.start:
            continue  # NOTE: Already a boundary — cutting here would only add an empty run
        split_wt_at(text_interval.t, off - text_interval.start)

    return ParagraphIndex(p_idx.p)


def hoist(el: E, ancestor: E) -> E:
    """Walk up from ``el`` until reaching the child of ``ancestor``."""
    while el is not None and el.getparent() is not ancestor:
        el = el.getparent()
    if el is None:
        raise StructureCorruptedError()  # ancestor was not on el's parent chain
    return el


def split_container(el: E, at_child: E) -> tuple[E, E]:
    """Split ``el`` before ``at_child``; both halves keep its tag and attributes."""
    right = etree.Element(el.tag)
    for k, v in el.attrib.items():
        right.set(k, v)
    el.addnext(right)
    for child in list(el)[el.index(at_child) :]:
        right.append(child)  # lxml moves rather than copies
    return el, right


def highlight_spans_xml(
    p_idx: ParagraphIndex, spans: list[tuple[int, int]], color: str = "yellow"
) -> ParagraphIndex:
    for a, b in spans:
        if b <= a:
            raise ValueError(f"Given span {a},{b} is invalid as b <= a")
    p_idx_cuts_applied = apply_cuts(p_idx, [x for span in spans for x in span])
    for interval in p_idx_cuts_applied.text_intervals:
        if any(a <= interval.start and interval.end <= b for a, b in spans):
            set_rpr_prop(interval.run, "w:highlight", color)
    return p_idx_cuts_applied


def add_comment_range_xml(
    p_idx: ParagraphIndex, start: int, end: int, cid: int
) -> ParagraphIndex:
    if end <= start:
        raise ValueError(f"Given span {start},{end} is invalid as end <= start")
    p_idx_cuts_applied = apply_cuts(p_idx, (start, end))
    first = p_idx_cuts_applied.interval_starting_at(start)
    last = p_idx_cuts_applied.interval_ending_at(end)
    if first is None or last is None:
        raise ValueError(f"Span {start},{end} is not anchorable in this paragraph")

    crs = range_marker("w:commentRangeStart", cid)
    cre = range_marker("w:commentRangeEnd", cid)

    first.anchor_before(crs)
    last.anchor_after(cre)
    # Word draws the anchor from the reference run, not from the range markers;
    # without it the comment exists but has nothing to click on.
    cre.addnext(comment_reference_run(cid))

    return p_idx_cuts_applied
