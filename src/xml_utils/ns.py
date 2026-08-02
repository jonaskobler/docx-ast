"""OOXML namespaces and the tag vocabulary every other module here speaks."""

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"

_NS = {"w": W, "r": R, "xml": XML}


def q(name: str) -> str:
    prefix, local = name.split(":")
    return f"{{{_NS[prefix]}}}{local}"


def ln(el) -> str:
    """Local name of an element's tag. Only valid for real elements — callers
    must skip comments and processing instructions first (see is_element)."""
    tag = el.tag
    return tag.rpartition("}")[2] if isinstance(tag, str) else ""


def is_element(el) -> bool:
    """False for lxml comment / processing-instruction nodes, whose .tag is a
    callable rather than a string."""
    return isinstance(el.tag, str)


# Run-level containers to descend into: both xml_edit's interval walk and the
# parser (docx_to_ast._parse_run_level) use this set, so the AST reports exactly
# the text the edit offsets are measured over. Paragraph.get_text() reads the
# interval walk directly, so the two cannot drift apart. w:sdt is handled
# separately because only its w:sdtContent is entered.
RUN_CONTAINERS = frozenset({"ins", "del", "moveFrom", "moveTo", "hyperlink"})

# Revision containers whose content the reader no longer sees: already deleted,
# or the source half of a move.
DELETED = ("del", "moveFrom")

# Revision containers holding content inserted but not yet accepted.
INSERTED = ("ins", "moveTo")
