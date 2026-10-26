"""Evidence integrity at the bridge and enforcement at the tool boundary (2026-10-23 audit, round 2)."""
from __future__ import annotations

import json
import os
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import bridge_client  # noqa: E402
import task_guard  # noqa: E402


# ------------------------------------------------------------------ bridge: shared result slots
class FakeVba:
    """Mimics WZTCBridge: runs a request file's ops in order; multi-row results go to a SMALL pool of slot files that
    are overwritten by later ops (the real defect). Rows of element E are E's own marker so misattribution is visible."""

    def __init__(self, tmp_path, pool=2):
        self.dir, self.pool, self.n = tmp_path, pool, 0
        self.req, self.resp = tmp_path / "request.tsv", tmp_path / "response.tsv"
        self.req.write_text("")
        self.resp.write_text("")

    def keyin(self, _cmd):
        out = []
        for ln in self.req.read_text().splitlines():
            parts = ln.split("\t")
            rid, op = parts[0], parts[1]
            kv = dict(p.split("=", 1) for p in parts[2:] if "=" in p)
            if op.startswith("GET_"):
                eid = int(kv["elementId"])
                slot = self.dir / f"slot{int(rid[1:]) % self.pool}.tsv"
                slot.write_text("vtxIndex\tx\ty\tz\n" + "".join(f"{i}\t{eid}\t{i}\t0\n" for i in range(eid % 5 + 2)))
                out.append(f"{rid}\tOK\trowCount={eid % 5 + 2}\tresultFile={slot}")
            else:
                slot = self.dir / f"slot{int(rid[1:]) % self.pool}.tsv"
                slot.write_text("a\tb\n1\t2\n")
                out.append(f"{rid}\tOK\trowCount=1\tresultFile={slot}")
        self.resp.write_text("\n".join(out) + "\n")


@pytest.fixture
def bridge(tmp_path, monkeypatch):
    vba = FakeVba(tmp_path)
    monkeypatch.setattr(bridge_client, "_send_keyin", vba.keyin)
    monkeypatch.setenv("QA_ALLOW_NO_TASK", "1")
    return bridge_client.Bridge(request_file=vba.req, response_file=vba.resp)


def test_batched_vertex_reads_come_back_attributed_to_the_right_element(bridge):
    ids = [11, 12, 13, 14, 15, 16]
    res = bridge._call_batch_impl([("GET_ELEMENT_VERTICES", {"elementId": e}) for e in ids])
    for e, r in zip(ids, res):
        assert [row["x"] for row in r["rows"]] == [str(e)] * (e % 5 + 2), (e, r["rows"])


def test_a_slot_collision_between_edit_ops_is_refused_not_guessed(bridge):
    ops = [("PLACE_TEXT_LABEL", {"x": 1}), ("PLACE_POLYLINE", {"x": 2}), ("PLACE_BLOCK", {"x": 3})]
    with pytest.raises(bridge_client.BridgeError, match="collision"):
        bridge._call_batch_impl(ops)


def test_shared_slot_detection():
    parsed = [{"resultFile": "A.tsv"}, {"resultFile": "a.tsv"}, {"resultFile": "B.tsv"}, {}]
    assert bridge_client._shared_result_slots(parsed) == {0, 1}


# ------------------------------------------------------------------ tool boundary
def _task(tmp_path, **kw):
    d = tmp_path / "T1"
    d.mkdir(exist_ok=True)
    t = {"id": "T1", "type": "cad.drafting", "status": "in_progress", "created": time.time()}
    t.update(kw)
    (d / "task.json").write_text(json.dumps(t))


@pytest.fixture
def guard(tmp_path, monkeypatch):
    monkeypatch.setattr(task_guard, "TASKS", tmp_path)
    monkeypatch.delenv("QA_ALLOW_NO_TASK", raising=False)
    task_guard._cache.update(at=0.0, open=None)
    return tmp_path


def test_reads_never_need_a_task(guard):
    task_guard.require_open_task(["GET_ELEMENT_VERTICES", "LIST_LEVELS", "DESCRIBE_DRAWING_STATE", "STATION_TO_POINT"])


def test_an_edit_without_an_open_task_is_refused(guard):
    with pytest.raises(task_guard.QAGateRequired, match="no open QA task"):
        task_guard.require_open_task(["GET_ELEMENT_VERTICES", "MOVE_ELEMENT"])


@pytest.mark.parametrize("op", ["PLACE_CELL", "DELETE_ELEMENT", "RUN_REGISTRY_COMMAND", "UNDO_LAST_OP", "SOME_NEW_OP"])
def test_every_non_read_op_counts_as_an_edit(guard, op):
    assert task_guard.is_edit(op)
    with pytest.raises(task_guard.QAGateRequired):
        task_guard.require_open_task([op])


def test_an_open_drafting_task_allows_edits(guard):
    _task(guard)
    task_guard.require_open_task(["MOVE_ELEMENT"])


@pytest.mark.parametrize("kw", [{"status": "closed"}, {"status": "blocked"}, {"status": "expired"}, {"type": "generic"},
                                {"created": time.time() - 13 * 3600}])
def test_closed_blocked_expired_or_wrong_type_tasks_do_not_count(guard, kw):
    _task(guard, **kw)
    with pytest.raises(task_guard.QAGateRequired):
        task_guard.require_open_task(["MOVE_ELEMENT"])


def test_maintenance_scripts_can_opt_out_explicitly(guard, monkeypatch):
    monkeypatch.setenv("QA_ALLOW_NO_TASK", "1")
    task_guard.require_open_task(["DELETE_ELEMENT"])


def test_the_bridge_itself_refuses_an_unguarded_edit(guard, tmp_path, monkeypatch):
    vba = FakeVba(tmp_path)
    sent = []
    monkeypatch.setattr(bridge_client, "_send_keyin", lambda c: sent.append(c))
    b = bridge_client.Bridge(request_file=vba.req, response_file=vba.resp)
    with pytest.raises(task_guard.QAGateRequired):
        b.call("MOVE_ELEMENT", elementId=1)
    assert not sent  # nothing reached MicroStation
