"""Shared XML element builders for revision tracking.

Centralises the repeated patterns for creating ``w:r``, ``w:del``, ``w:ins``
elements and the ``before / del / ins / after`` splice used by both
``RevisionParagraph`` and ``RevisionRun``.
"""

from __future__ import annotations

import contextlib
from typing import Callable

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from lxml import etree


def revision_attrs(rev_id: int, author: str, now: str) -> dict[str, str]:
    """Build the standard ``{w:id, w:author, w:date}`` attribute dict."""
    return {qn("w:id"): str(rev_id), qn("w:author"): author, qn("w:date"): now}


def make_text_run(text: str) -> OxmlElement:
    """Create a ``<w:r><w:t>text</w:t></w:r>`` element with space preservation."""
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def make_del_element(deleted_text: str, author: str, rev_id: int, now: str) -> OxmlElement:
    """Create a ``<w:del><w:r><w:delText>text</w:delText></w:r></w:del>`` element."""
    del_elem = OxmlElement("w:del", attrs=revision_attrs(rev_id, author, now))
    del_r = OxmlElement("w:r")
    del_text_elem = OxmlElement("w:delText")
    del_text_elem.text = deleted_text
    del_r.append(del_text_elem)
    del_elem.append(del_r)
    return del_elem


def make_ins_element(insert_text: str, author: str, rev_id: int, now: str) -> OxmlElement:
    """Create a ``<w:ins><w:r><w:t>text</w:t></w:r></w:ins>`` element."""
    ins_elem = OxmlElement("w:ins", attrs=revision_attrs(rev_id, author, now))
    ins_r = OxmlElement("w:r")
    ins_t = OxmlElement("w:t")
    ins_t.text = insert_text
    ins_r.append(ins_t)
    ins_elem.append(ins_r)
    return ins_elem


def make_comment_range_start(comment_id: int) -> OxmlElement:
    """Create a ``<w:commentRangeStart w:id="N"/>`` marker element."""
    return OxmlElement("w:commentRangeStart", attrs={qn("w:id"): str(comment_id)})


def make_comment_range_end(comment_id: int) -> OxmlElement:
    """Create a ``<w:commentRangeEnd w:id="N"/>`` marker element."""
    return OxmlElement("w:commentRangeEnd", attrs={qn("w:id"): str(comment_id)})


def make_comment_reference_run(comment_id: int) -> OxmlElement:
    """Create a ``<w:r>`` wrapping a ``<w:commentReference w:id="N"/>`` element.

    This is the "anchor" run Word reads when navigating to the commented range.
    """
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "CommentReference")
    rPr.append(rStyle)
    r.append(rPr)
    r.append(OxmlElement("w:commentReference", attrs={qn("w:id"): str(comment_id)}))
    return r


def wrap_with_comment(first_elem: etree._Element, last_elem: etree._Element, comment_id: int) -> None:
    """Wrap the sibling range *first_elem* … *last_elem* with comment range markers.

    Inserts ``w:commentRangeStart`` immediately before *first_elem* and
    ``w:commentRangeEnd`` followed by a ``w:r/w:commentReference`` immediately
    after *last_elem*.  Both elements must share the same parent.
    """
    parent = first_elem.getparent()
    if parent is None or last_elem.getparent() is not parent:
        raise ValueError("first_elem and last_elem must share the same parent")

    first_idx = list(parent).index(first_elem)
    parent.insert(first_idx, make_comment_range_start(comment_id))

    # last_elem's position shifted by 1 because of the insertion before first_elem
    last_idx = list(parent).index(last_elem)
    parent.insert(last_idx + 1, make_comment_range_end(comment_id))
    parent.insert(last_idx + 2, make_comment_reference_run(comment_id))


def splice_tracked_replace(
    parent: etree._Element,
    index: int,
    before_text: str | None,
    deleted_text: str,
    insert_text: str,
    after_text: str | None,
    author: str,
    next_id_fn: Callable[[], int],
    now: str,
) -> tuple[etree._Element, etree._Element]:
    """Insert the before-run / w:del / w:ins / after-run sequence into *parent* at *index*.

    Returns:
        The ``(w:del, w:ins)`` element pair that was inserted, so the caller
        can attach further metadata (e.g. comment range markers) around them.
    """
    insert_idx = index
    if before_text:
        parent.insert(insert_idx, make_text_run(before_text))
        insert_idx += 1

    del_elem = make_del_element(deleted_text, author, next_id_fn(), now)
    parent.insert(insert_idx, del_elem)
    insert_idx += 1

    ins_elem = make_ins_element(insert_text, author, next_id_fn(), now)
    parent.insert(insert_idx, ins_elem)
    insert_idx += 1

    if after_text:
        parent.insert(insert_idx, make_text_run(after_text))

    return del_elem, ins_elem


def next_revision_id(element: etree._Element) -> int:
    """Generate the next unique revision ID by scanning the document tree from *element*."""
    max_id = 0
    for ins_or_del in element.xpath("//w:ins | //w:del"):
        id_val = ins_or_del.get(qn("w:id"))
        if id_val is not None:
            with contextlib.suppress(ValueError):
                max_id = max(max_id, int(id_val))
    return max_id + 1
