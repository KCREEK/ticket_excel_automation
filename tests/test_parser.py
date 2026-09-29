import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ticket_tool.parser import parse_file, parse_text  # noqa: E402

SAMPLE = os.path.join(os.path.dirname(__file__), "..", "samples", "sample_notes.txt")


def rows():
    return parse_file(SAMPLE).rows


def test_every_entry_is_kept_including_repeated_ids():
    r = rows()
    assert len(r) == 10                       # nothing dropped
    ids = [x.ticket_id for x in r]
    assert ids.count("T20260101.0003") == 2   # same ticket, two comments
    assert ids.count("T20260101.0001") == 2


def test_header_fields():
    r = rows()[0]
    assert (r.ticket_id, r.priority, r.status) == ("T20260101.0001", "P3", "In Progress")
    assert r.title == "Patching Review - HOST-0001"
    assert "Failed Actions: 1" in r.description


def test_title_with_dashes_and_double_spaces():
    r = rows()[1]
    assert r.title == "SOC Alert : Example Threat- 3 - Moderate - [CS0000001] - Example Corp"
    assert r.priority == "P3" and r.status == "Assigned"


def test_boilerplate_stripping_toggle():
    on = parse_file(SAMPLE, strip_boilerplate=True).rows[1].description
    off = parse_file(SAMPLE, strip_boilerplate=False).rows[1].description
    assert on == "We contacted the customer to review; no failures in the last 12 hours."
    assert off.startswith("Recent Comments") and "Jane Doe" in off


def test_grouped_headers_share_body_and_inherit_priority_status():
    g = rows()[4:7]
    assert [x.ticket_id for x in g] == ["T20260101.0004", "T20260101.0005", "T20260101.0006"]
    assert {x.priority for x in g} == {"P2"}
    assert {x.status for x in g} == {"Complete"}
    assert len({x.description for x in g}) == 1
    assert g[0].description.startswith("Device transitioned")


def test_priority_filled_from_same_ticket_only():
    text = ("1. T20260101.0010 - A | P1 | Assigned\n=====\nx\n_____\n"
            "2. T20260101.0010 - A\n=====\ny\n_____\n"
            "3. T20260101.0011 - B\n=====\nz\n_____\n")
    r = parse_text(text).rows
    assert r[1].priority == "P1" and r[1].status == ""   # status is NOT inherited
    assert r[2].priority == ""


def test_entry_without_number_or_separator_at_end_of_file():
    last = rows()[-1]
    assert last.ticket_id == "T20260101.0008"
    assert last.priority == "" and last.status == ""
    assert last.description.startswith("Reviewed the alert")
    assert "1. Executive Summary" in last.description      # numbered body lines are not headers
    assert not last.description.endswith("\n")


def test_day_marker_ignored_and_no_stray_lines():
    res = parse_file(SAMPLE)
    assert res.ignored_lines == []
    assert rows()[8].ticket_id == "T20260101.0007"


def test_crlf_and_utf16(tmp_path):
    text = open(SAMPLE, encoding="utf-8").read().replace("\n", "\r\n")
    p = tmp_path / "n.txt"
    p.write_bytes(text.encode("utf-16"))
    assert len(parse_file(str(p)).rows) == 10
