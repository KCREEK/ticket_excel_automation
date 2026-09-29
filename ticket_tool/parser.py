"""Parser for the daily rough ticket-notes .txt file.

Patterns observed in the notes file
-----------------------------------
* Header line:   ``[N.] T20260706.0190 - Title text | P3 | In Progress``
  - the ``N.`` numbering is optional
  - ``| P# | Status`` is optional (some entries have neither)
  - the title itself may contain " - " (e.g. "... - Moderate - [CS6458323] - Customer")
* ``=====`` underline after the header, then the body (the "action taken").
* Body ends at a ``_____`` line, at the next header, or at end of file.
* Several consecutive header lines with no body between them share the one
  body that follows (e.g. ten "Disk ... is Failed" alerts closed with one
  comment). Priority/status written on the first header of the group are
  inherited by the others.
* ``_____XXXXX_____`` is a visual day/shift divider - ignored.
* The same ticket ID can appear many times - every occurrence is kept.
  Nothing in this module removes or merges rows.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List

TICKET_ID_PATTERN = r"T\d{8}\.\d{3,5}"

HEADER_RE = re.compile(
    rf"^\s*(?:(?P<num>\d+)\s*[.)]\s*)?(?P<id>{TICKET_ID_PATTERN})"
    r"\s+[-\u2013\u2014]\s+(?P<rest>\S.*?)\s*$"
)
TAIL_RE = re.compile(
    r"\s*\|\s*(?P<prio>P[0-5])\s*(?:\|\s*(?P<status>[^|]*?))?\s*$", re.IGNORECASE
)
EQ_RE = re.compile(r"^\s*={5,}\s*$")
END_RE = re.compile(r"^\s*_{5,}\s*$")
DAY_MARKER_RE = re.compile(r"^\s*_{3,}\s*[Xx]{3,}\s*_{3,}\s*$")
# "2026-09-29 12:13:57 EDT - Bailey Stiles Additional comments"
COMMENT_META_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}\s+[A-Za-z]{2,5}\s+-\s+.+$"
)


@dataclass
class TicketRow:
    ticket_id: str
    priority: str = ""
    status: str = ""
    title: str = ""
    description: str = ""
    time_taken: str = ""
    source_no: str = ""

    def as_list(self) -> List[str]:
        return [self.ticket_id, self.priority, self.status, self.title,
                self.description, self.time_taken]


@dataclass
class ParseResult:
    rows: List[TicketRow] = field(default_factory=list)
    ignored_lines: List[str] = field(default_factory=list)


def read_text_file(path: str) -> str:
    """Read a Notepad file whatever encoding Notepad saved it in."""
    with open(path, "rb") as fh:
        raw = fh.read()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace")


def _clean_body(lines: List[str], strip_boilerplate: bool) -> str:
    lines = [ln.rstrip() for ln in lines]
    first = next((ln.strip() for ln in lines if ln.strip()), "")
    if strip_boilerplate and first.lower() == "recent comments":
        out, heading_removed = [], False
        for ln in lines:
            s = ln.strip()
            if not heading_removed and s.lower() == "recent comments":
                heading_removed = True
                continue
            if COMMENT_META_RE.match(s):
                continue
            out.append(ln)
        lines = out
    cleaned: List[str] = []
    for ln in lines:                      # collapse runs of blank lines
        if not ln.strip():
            if cleaned and cleaned[-1] != "":
                cleaned.append("")
        else:
            cleaned.append(ln)
    while cleaned and cleaned[-1] == "":
        cleaned.pop()
    return "\n".join(cleaned)


def _split_header(m):
    rest = m.group("rest")
    prio = status = ""
    tail = TAIL_RE.search(rest)
    if tail:
        prio = tail.group("prio").upper()
        status = (tail.group("status") or "").strip()
        rest = rest[: tail.start()]
    title = re.sub(r"\s+", " ", rest).strip()
    return m.group("num") or "", m.group("id"), title, prio, status


def parse_text(text: str, strip_boilerplate: bool = True) -> ParseResult:
    result = ParseResult()
    group: List[tuple] = []
    body: List[str] = []
    state = "seek"          # seek -> header -> body

    def finish():
        nonlocal group, body
        if group:
            desc = _clean_body(body, strip_boilerplate)
            prio = next((g[3] for g in group if g[3]), "")
            status = next((g[4] for g in group if g[4]), "")
            for num, tid, title, p, s in group:
                result.rows.append(TicketRow(tid, p or prio, s or status,
                                             title, desc, "", num))
        group, body = [], []

    for line in text.splitlines():
        m = HEADER_RE.match(line)
        if m:
            if state == "body":
                finish()
            group.append(_split_header(m))
            state = "header"
        elif END_RE.match(line) or DAY_MARKER_RE.match(line):
            finish()
            state = "seek"
        elif EQ_RE.match(line):
            if state == "header":
                state = "body"
        elif state == "header":
            if line.strip():            # body with no ===== line (last entry)
                state = "body"
                body.append(line)
        elif state == "body":
            body.append(line)
        elif line.strip():
            result.ignored_lines.append(line.strip())
    finish()

    # Priority never changes between updates of a ticket, so fill blanks from
    # another entry of the same ticket. Rows are never removed or merged.
    known: Dict[str, str] = {}
    for r in result.rows:
        if r.priority and r.ticket_id not in known:
            known[r.ticket_id] = r.priority
    for r in result.rows:
        if not r.priority:
            r.priority = known.get(r.ticket_id, "")
    return result


def parse_file(path: str, strip_boilerplate: bool = True) -> ParseResult:
    return parse_text(read_text_file(path), strip_boilerplate)
