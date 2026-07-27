from __future__ import annotations

from diff_match_patch import diff_match_patch

from docx_ast.nodes import (
    Del,
    DelText,
    Ins,
    MoveFrom,
    MoveTo,
    Node,
    Paragraph,
    Text,
    link_parents,
    walk,
)


def _root(node: Node) -> Node:
    while node.parent is not None:
        node = node.parent
    return node


def _next_rev_id(root: Node) -> int:
    # NOTE: A bit silly to do this on every call, but I am lazy for now
    max_id = 0

    def walk(n: Node) -> None:
        nonlocal max_id
        if isinstance(n, (Ins, Del, MoveFrom, MoveTo)):
            max_id = max(max_id, n.id)
        for child in getattr(n, "children", None) or []:
            walk(child)
        footnotes = getattr(n, "footnotes", None)
        if footnotes:
            for fn in footnotes.values():
                walk(fn)

    walk(root)
    return max_id + 1


def alter_text_tracked(
    node: Text,
    new_text: str,
    *,
    author: str = "",
    date: str | None = None,
) -> list[Node]:
    parent = node.parent
    siblings = getattr(parent, "children", None)
    if parent is None or siblings is None:
        raise ValueError(
            "node has no linked parent with children; call link_parents(root) first"
        )

    index = next((i for i, s in enumerate(siblings) if s is node), None)
    if index is None:
        raise ValueError(
            "node is not among its parent's children (stale parent pointer?)"
        )

    dmp = diff_match_patch()
    diffs = dmp.diff_main(node.text, new_text)
    dmp.diff_cleanupSemantic(diffs)

    rev_id = _next_rev_id(_root(node))
    style_id = node.style_id

    replacement: list[Node] = []
    for op, chunk in diffs:
        if not chunk:
            continue
        if op == dmp.DIFF_EQUAL:
            replacement.append(Text(chunk, style_id=style_id))
        elif op == dmp.DIFF_DELETE:
            replacement.append(
                Del(
                    id=rev_id,
                    author=author,
                    date=date,
                    children=[DelText(chunk, style_id=style_id)],
                )
            )
            rev_id += 1
        elif op == dmp.DIFF_INSERT:
            replacement.append(
                Ins(
                    id=rev_id,
                    author=author,
                    date=date,
                    children=[Text(chunk, style_id=style_id)],
                )
            )
            rev_id += 1

    siblings[index : index + 1] = replacement
    for new_node in replacement:
        link_parents(new_node, parent)
    return replacement


def alter_paragraph_tracked(
    para: Paragraph,
    new_text: str,
    *,
    author: str = "",
    date: str | None = None,
) -> list[Node]:
    text_nodes = [n for n in walk(para) if isinstance(n, Text) and n.text]
    rev_id = _next_rev_id(_root(para))

    if not text_nodes:
        if new_text:
            para.children.append(
                Ins(id=rev_id, author=author, date=date, children=[Text(new_text)])
            )
            link_parents(para, para.parent)
        return para.children

    full = "".join(t.text for t in text_nodes)
    starts: list[int] = []
    ends: list[int] = []
    pos = 0
    for t in text_nodes:
        starts.append(pos)
        pos += len(t.text)
        ends.append(pos)

    def locate(p: int) -> int:
        for i in range(len(text_nodes)):
            if starts[i] <= p < ends[i]:
                return i
        return len(text_nodes) - 1

    def choose_insert(p: int) -> int:
        for i in range(len(text_nodes)):
            if starts[i] < p < ends[i]:
                return i
        left = next((i for i in range(len(text_nodes)) if ends[i] == p), None)
        right = next((i for i in range(len(text_nodes)) if starts[i] == p), None)
        if left is None:
            return right
        if right is None:
            return left
        lc = full[p - 1]
        rc = full[p]
        if lc.isspace() and not rc.isspace():
            return right
        if rc.isspace() and not lc.isspace():
            return left
        return left

    dmp = diff_match_patch()
    diffs = dmp.diff_main(full, new_text)
    dmp.diff_cleanupSemantic(diffs)

    expansions: dict[int, list[Node]] = {id(t): [] for t in text_nodes}

    cursor = 0
    for op, chunk in diffs:
        if not chunk:
            continue
        if op == dmp.DIFF_INSERT:
            t = text_nodes[choose_insert(cursor)]
            expansions[id(t)].append(
                Ins(
                    id=rev_id,
                    author=author,
                    date=date,
                    children=[Text(chunk, style_id=t.style_id)],
                )
            )
            rev_id += 1
            continue

        keep = op == dmp.DIFF_EQUAL
        del_id = None
        if not keep:
            del_id = rev_id
            rev_id += 1
        end = cursor + len(chunk)
        cur = cursor
        while cur < end:
            i = locate(cur)
            seg_end = min(end, ends[i])
            piece = full[cur:seg_end]
            t = text_nodes[i]
            if keep:
                expansions[id(t)].append(Text(piece, style_id=t.style_id))
            else:
                expansions[id(t)].append(
                    Del(
                        id=del_id,
                        author=author,
                        date=date,
                        children=[DelText(piece, style_id=t.style_id)],
                    )
                )
            cur = seg_end
        cursor = end

    for t in text_nodes:
        kids = t.parent.children
        idx = next(i for i, c in enumerate(kids) if c is t)
        kids[idx : idx + 1] = expansions[id(t)]
    link_parents(para, para.parent)
    return para.children
