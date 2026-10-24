"""Help examples, op lines and view notes, carried over from the Word editor.

  * every op's help example is a curated call, never a schema placeholder
  * an op's line names its actions (manage_comment and manage_chart were
    cut to a sentence that named none) and stays one line
  * get_grid_view says what its table leaves out: notes and threaded
    comments on the shown cells, the sheet's charts and images
  * get_server_info reports approx_active_tokens, the cost of the tool list
  * xl_read lists the cheap orientation reads first
  * results name a cell at_cell or top_left; "anchor" is only a
    get_grid_view token; a chart list speaks create's chart_type
"""

from __future__ import annotations

import json

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp import registry
from xlsx_mcp import server as srv
from xlsx_mcp.settings import Settings
from xlsx_mcp.surface import examples, plan
from xlsx_mcp.surface.plan import op_text


def _call(tool, op, **args):
    payload, err, _ = fh.start().call(tool, {"op": op, "args": args})
    assert not err, payload
    return payload


def test_e1_every_op_has_a_curated_example():
    missing = sorted(set(registry.public_ops()) - set(examples.EXAMPLES))
    assert missing == []
    for op, ex in examples.EXAMPLES.items():
        assert "text" not in [v for v in ex.values() if isinstance(v, str)], op


def test_e1_help_serves_the_curated_example():
    payload, err, _ = fh.start().call("xl_help", {"op": "audit_styles"})
    assert not err and payload["example"] == {"op": "audit_styles",
                                              "args": {"path": "workbook.xlsx", "top": 5}}


@pytest.mark.parametrize("op,words", [
    ("manage_comment", ("add", "edit", "delete", "list")),
    ("manage_chart", ("create", "list", "delete")),
    ("set_merge", ("merge", "unmerge", "list")),
    ("com_set_sparkline", ("create", "clear", "list")),
])
def test_e4_an_action_ops_line_names_its_actions(op, words):
    line = op_text(registry.public_ops()[op]).lower()
    assert all(w in line for w in words), line


def test_e4_every_op_line_stays_one_line():
    long = {n: len(op_text(s)) for n, s in registry.public_ops().items() if len(op_text(s)) > 200}
    assert long == {}


def _decorated(tmp_path):
    p = str(tmp_path / "g.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Month", "Sales"])
    for m, v in (("Jan", 3), ("Feb", 5), ("Mar", 4)):
        ws.append([m, v])
    wb.save(p)
    _call("xl_format", "manage_comment", path=p, action="add", location="B2", text="check")
    _call("xl_format", "manage_chart", path=p, action="create", chart_type="line",
          data="B1:B4", categories="A2:A4", at_cell="E2")
    return p


def test_e5_the_view_names_what_it_leaves_out(tmp_path):
    p = _decorated(tmp_path)
    view = _call("xl_read", "get_grid_view", path=p)
    hidden = view["not_shown"]
    assert hidden["comments"]["cells"] == ["B2"]
    assert hidden["charts"]["items"] == [{"chart_type": "line", "at_cell": "E2"}]
    assert "manage_chart" in hidden["charts"]["read_with"]
    # A view of cells with no note does not list the note.
    part = _call("xl_read", "get_grid_view", path=p, location="A1:A4")
    assert "comments" not in part["not_shown"]


def test_e5_a_plain_sheet_has_nothing_to_name(tmp_path):
    p = str(tmp_path / "p.xlsx")
    srv.create_workbook(p)
    srv.set_cell(p, "A1", 1)
    assert "not_shown" not in _call("xl_read", "get_grid_view", path=p)


def test_e5_a_threaded_comment_is_named_too(tmp_path):
    from test_shared_word_editor_behaviours import _with_thread
    p = _with_thread(tmp_path / "t.xlsx")
    view = _call("xl_read", "get_grid_view", path=p)
    assert view["not_shown"]["comments"]["cells"] == ["B8"]


def test_e7_the_token_bill_matches_the_wire_within_two_percent():
    res = fh.stdio_session([("xl_read", {"op": "get_server_info", "args": {}})])
    wire = sum(round(len(json.dumps(t, ensure_ascii=False, separators=(",", ":"))) / 4)
               for t in res["tools"])
    reported = json.loads(res["results"][0]["text"])["surface"]["approx_active_tokens"]
    assert abs(reported - wire) <= 0.02 * wire, (reported, wire)


def test_c5_xl_read_leads_with_the_cheap_orientation_reads():
    tool = plan.build(Settings()).tools["xl_read"]
    assert tool.ops[:3] == ["get_workbook_metadata", "get_grid_view", "read_range"]
    desc = tool.description
    assert desc.index("get_workbook_metadata") < desc.index("audit_formulas")


def test_section3_results_name_cells_without_anchor(tmp_path):
    p = _decorated(tmp_path)
    out = _call("xl_edit", "write_range", path=p, location="G1", data=[[1]])
    assert out["changed"]["top_left"] == "G1" and "anchor" not in out["changed"]
    out = _call("xl_edit", "import_data", path=p, source="a,b\n1,2\n", location="J1")
    assert out["changed"]["imported"]["top_left"] == "J1"
    (ch,) = _call("xl_format", "manage_chart", path=p, action="list")["charts"]
    assert ch["chart_type"] == "line" and ch["at_cell"] == "E2" and "anchor" not in ch
