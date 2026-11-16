"""Notes (legacy comments), threaded comments and the people who wrote them.

A note's box (its position, size and look) is in the sheet's VML drawing (:class:`xlio.Vml`),
matched to the note by its cell (``x:Row``, ``x:Column``). Excel for Mac in Office 2021 shows
threaded comments as threads with their own mark; a threaded comment also has a placeholder
note, whose VML ``x:Visible`` Excel ignores.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Mapping
from dataclasses import dataclass, field

from xlio._rich import RichString, rich
from xlio._types import flag, integer, local

__all__ = ["Comment", "Person", "ThreadedComment", "read_comments", "read_persons", "read_threaded"]


@dataclass(frozen=True, slots=True)
class Comment:
    """A note on a cell.

    Attributes:
        ref: The cell.
        author: The author's name, or ``None`` when ``author_id`` leads nowhere.
        author_id: The index into the part's authors.
        text: The text with its runs.
        uid: ``xr:uid``, which a threaded comment's placeholder note shares, or ``None``.
        source: The element as written.
    """

    ref: str
    author: str | None
    author_id: int | None
    text: RichString
    uid: str | None = None
    source: ET.Element | None = field(default=None, compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class Person:
    """A person who wrote threaded comments (``persons`` part)."""

    id: str
    display_name: str
    user_id: str | None = None
    provider_id: str | None = None


@dataclass(frozen=True, slots=True)
class ThreadedComment:
    """A comment of a thread.

    Attributes:
        ref: The cell.
        id: Its id.
        parent_id: The thread's first comment, for a reply; ``None`` for the first.
        person_id: Its author's id.
        person: Its author, or ``None`` when the id leads nowhere.
        date: The time written (``dT``), as written, or ``None``.
        done: The thread is resolved (on the first comment).
        text: The text.
        source: The element as written.
    """

    ref: str
    id: str | None
    parent_id: str | None
    person_id: str | None
    person: Person | None
    date: str | None
    done: bool
    text: str
    source: ET.Element | None = field(default=None, compare=False, repr=False)


def _uid(element: ET.Element) -> str | None:
    return next((v for k, v in element.attrib.items() if local(k) == "uid"), None)


def read_comments(root: ET.Element) -> tuple[Comment, ...]:
    """A comments part's notes, in file order."""
    authors: list[str] = []
    out = []
    for child in root:
        name = local(child.tag)
        if name == "authors":
            authors = [a.text or "" for a in child if local(a.tag) == "author"]
        elif name == "commentList":
            for c in child:
                if local(c.tag) != "comment":
                    continue
                author_id = integer(c.get("authorId"))
                text = next((t for t in c if local(t.tag) == "text"), None)
                out.append(
                    Comment(
                        c.get("ref", ""),
                        authors[author_id]
                        if author_id is not None and 0 <= author_id < len(authors)
                        else None,
                        author_id,
                        rich(text) if text is not None else RichString(""),
                        _uid(c),
                        c,
                    )
                )
    return tuple(out)


def read_persons(root: ET.Element) -> dict[str, Person]:
    """A persons part, by id."""
    return {
        p.get("id", ""): Person(
            p.get("id", ""), p.get("displayName", ""), p.get("userId"), p.get("providerId")
        )
        for p in root
        if local(p.tag) == "person"
    }


def read_threaded(root: ET.Element, persons: Mapping[str, Person]) -> tuple[ThreadedComment, ...]:
    """A threaded comments part's comments, in file order."""
    out = []
    for c in root:
        if local(c.tag) != "threadedComment":
            continue
        text = next((t.text or "" for t in c if local(t.tag) == "text"), "")
        person_id = c.get("personId")
        out.append(
            ThreadedComment(
                c.get("ref", ""),
                c.get("id"),
                c.get("parentId"),
                person_id,
                persons.get(person_id) if person_id is not None else None,
                c.get("dT"),
                flag(c.get("done"), False),
                text,
                c,
            )
        )
    return tuple(out)
