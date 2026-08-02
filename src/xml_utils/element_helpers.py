from typing import Literal

from lxml import etree
from lxml.etree import _Element as E

from xml_utils.exceptions import StructureCorruptedError
from xml_utils.ns import R, W, is_element, ln, q

# CT_RPr's child order (ECMA-376 §17.3.2.27). Word tolerates a wrong order but
# strict validators do not, so properties are inserted at the right index.
# Lets just trust Claude
_RPR_ORDER = (
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
)


def comment_reference_run(cid: int) -> E:
    run = etree.Element(q("w:r"))
    rPr = etree.SubElement(run, q("w:rPr"))
    etree.SubElement(rPr, q("w:rStyle")).set(q("w:val"), "CommentReference")
    etree.SubElement(run, q("w:commentReference")).set(q("w:id"), str(cid))
    return run


def range_marker(
    tag: Literal["w:commentRangeStart", "w:commentRangeEnd"], cid: int
) -> E:
    el = etree.Element(q(tag))
    el.set(q("w:id"), str(cid))
    return el


def comments_root() -> E:
    """An empty ``word/comments.xml`` root, for a document that had none."""
    return etree.Element(q("w:comments"), nsmap={"w": W, "r": R})


def comment_element(
    cid: int,
    text: str,
    *,
    author: str = "",
    initials: str = "",
    date: str | None = None,
) -> E:
    """A ``w:comment`` whose body is one plain-text paragraph.

    Appended to the comments part rather than rebuilt from it, so comments that
    were only ever read keep whatever they came with — replies, formatting,
    paraId, resolved state — none of which the AST models.
    """
    c = etree.Element(q("w:comment"))
    c.set(q("w:id"), str(cid))
    c.set(q("w:author"), author)
    if initials:
        c.set(q("w:initials"), initials)
    if date is not None:
        c.set(q("w:date"), date)
    p = etree.SubElement(c, q("w:p"))
    p.append(new_text_run(text))
    return c


def preserve_space(t_el: E) -> None:
    t_el.set(q("xml:space"), "preserve")


def new_text_run(text: str, rPr: E | None = None, *, tag: str = "w:t") -> E:
    run = etree.Element(q("w:r"))
    if rPr is not None:
        run.append(rPr)
    t = etree.SubElement(run, q(tag))
    preserve_space(t)
    t.text = text
    return run


def get_or_insert_rpr(run: E) -> E:
    rPr = run.find(q("w:rPr"))
    if rPr is None:
        rPr = etree.Element(q("w:rPr"))
        run.insert(0, rPr)
    return rPr


def _rank(local: str) -> int:
    """Position of a property in CT_RPr's child order; unknown ones sort last."""
    return _RPR_ORDER.index(local) if local in _RPR_ORDER else len(_RPR_ORDER)


def set_rpr_prop(target: E, name: str, val: str | None = None) -> E:
    """Set property ``name`` on a run's ``w:rPr``, or on a bare ``w:rPr``."""
    if ln(target) == "rPr":
        rPr = target
    elif ln(target) == "r":
        rPr = get_or_insert_rpr(target)
    else:
        raise ValueError("Trying to set a rPr on neither a run nor a rPr element")
    el = rPr.find(q(name))
    if el is None:
        el = etree.Element(q(name))
        rank = _rank(name.split(":")[1])
        pos = len(rPr)
        for i, existing in enumerate(rPr):
            if _rank(ln(existing)) > rank:
                pos = i
                break
        rPr.insert(pos, el)
    if val is not None:
        el.set(q("w:val"), val)
    return el


def wrap_siblings(first: E, last: E, wrapper: E) -> E:
    parent = first.getparent()
    if parent is None:
        raise ValueError("Trying to wrap a root")
    start, stop = parent.index(first), parent.index(last)
    if stop < start:
        raise ValueError("last comes before first")
    span = parent[start : stop + 1]  # NOTE: before the insert shifts everything right
    parent.insert(start, wrapper)
    for el in span:
        if not is_element(el):
            continue
        tag = ln(el)
        if tag == "proofErr":
            parent.remove(el)
        elif tag in (
            "bookmarkStart",
            "bookmarkEnd",
        ):  # NOTE: Claude suggested to move this to the front, not sure about it
            wrapper.addprevious(el)
        else:
            wrapper.append(el)  # NOTE: append moves the element
    return wrapper


def nearest_common_ancestor(a: E, b: E) -> E | None:
    seen = set()
    node = a
    while node is not None:
        seen.add(id(node))
        node = node.getparent()
    node = b
    while node is not None:
        if id(node) in seen:
            return node
        node = node.getparent()
    return None


def hoist(el: E, ancestor: E) -> E:
    """Walk up from ``el`` until reaching the child of ``ancestor``."""
    while el is not None and el.getparent() is not ancestor:
        parent = el.getparent()
        if parent is None:
            raise StructureCorruptedError
        el = parent
    return el
