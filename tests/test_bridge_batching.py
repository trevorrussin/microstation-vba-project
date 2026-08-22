"""Build speed: batch bridge round trips, without changing what gets placed.

Measured live 2026-08-20, per bridge round trip on this install:
    ms_connect.get_microstation_app()   ~680 ms   <-- paid on EVERY call
      of which app.VBE.VBProjects()     ~500 ms
    app.CadInputQueue.SendKeyin()       ~100 ms
A 309-element road took ~1248 s (~4.0 s/element) because striping ran THREE
round trips per segment (place -> level -> symbology) and each one re-proved
which MicroStation to talk to.

Two fixes, both verified live: a per-thread COM handle cache (780 ms -> 113 ms
per call) and batching those three passes across segments via
bridge_client.call_batch, which VBA already supported (RunChatToolRequest
loops every request line). Result: ~183 ms/element, ~22x faster, identical
geometry/levels/colours.

These tests pin the round-trip COUNT, since a refactor back to per-element
calls would still pass every geometry test while quietly restoring the
20-minute build.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops as ops  # noqa: E402


class _RecordingBridge:
    """Stands in for the real bridge, counting keyin round trips."""

    def __init__(self, fail_ops: set[str] | None = None):
        self.batches: list[list[tuple[str, dict]]] = []
        self.fail_ops = fail_ops or set()
        self._next_id = 900000

    def call_batch(self, ops_list):
        self.batches.append(list(ops_list))
        out = []
        for op, _params in ops_list:
            if op in self.fail_ops:
                out.append({"status": "ERROR", "note": f"{op} refused"})
                continue
            self._next_id += 1
            out.append({"status": "OK", "elementId": str(self._next_id),
                        "deleted": 1})
        return out

    def call(self, op, **params):
        return self.call_batch([(op, params)])[0]

    @property
    def round_trips(self) -> int:
        return len(self.batches)

    def ops_of(self, name: str) -> list[dict]:
        return [p for b in self.batches for (o, p) in b if o == name]


@pytest.fixture
def rec(monkeypatch):
    b = _RecordingBridge()
    monkeypatch.setattr(ops, "_bridge", b)
    return b


def _segs(n: int, kinds=("lane",)) -> list[dict]:
    out = []
    for i in range(n):
        out.append({
            "x1": float(i * 10), "y1": 0.0, "x2": float(i * 10 + 8), "y2": 0.0,
            "style": "dash", "kind": kinds[i % len(kinds)], "row": i,
        })
    return out


def test_striping_uses_a_handful_of_round_trips_not_three_per_segment(rec):
    placed, errors, _ = ops._place_road_line_segments(
        _segs(60), reason_prefix="test", need_yellow=False)
    assert errors == []
    assert len(placed) == 60
    # Old behaviour: 3 * 60 = 180 round trips. Batched: 3 passes / 60-op chunks.
    assert rec.round_trips == 3, f"expected 3 batched round trips, got {rec.round_trips}"
    assert len(rec.ops_of("PLACE_POLYLINE")) == 60
    assert len(rec.ops_of("CHANGE_ELEMENT_LEVEL")) == 60
    assert len(rec.ops_of("CHANGE_ELEMENT_SYMBOLOGY")) == 60


def test_large_road_is_chunked_not_one_giant_keyin(rec):
    ops._place_road_line_segments(_segs(130), reason_prefix="test", need_yellow=False)
    # 130 ops -> ceil(130/60)=3 chunks per pass, 3 passes = 9 round trips.
    # Still ~43x fewer than the 390 the per-element version would have made.
    assert rec.round_trips == 9
    assert all(len(b) <= ops._BRIDGE_BATCH_SIZE for b in rec.batches)


def test_every_segment_still_gets_level_and_symbology(rec):
    ops._place_road_line_segments(_segs(5), reason_prefix="test", need_yellow=False)
    lvl = rec.ops_of("CHANGE_ELEMENT_LEVEL")
    sym = rec.ops_of("CHANGE_ELEMENT_SYMBOLOGY")
    assert all(p["level"] == "Default" for p in lvl)
    assert all(p["ownElementOnly"] == "Y" for p in lvl + sym)
    assert all(p["weight"] == 0 for p in sym)
    # Element IDs from the place pass must flow into the follow-up passes.
    placed_ids = {p["elementId"] for p in rec.ops_of("PLACE_POLYLINE")
                  if "elementId" in p}
    assert not placed_ids, "place op should not carry an elementId"
    assert {p["elementId"] for p in lvl} == {p["elementId"] for p in sym}


def test_geometry_and_ordering_are_unchanged(rec):
    segs = _segs(4)
    placed, _, _ = ops._place_road_line_segments(
        segs, reason_prefix="test", need_yellow=False)
    # Same order, same endpoints as the input segments.
    for seg, out in zip(segs, placed):
        assert out["row"] == seg["row"]
        assert abs(out["x1"] - seg["x1"]) < 1e-9
        assert abs(out["x2"] - seg["x2"]) < 1e-9
        assert out["vertexCount"] == 2


def test_curved_segments_keep_all_their_vertices(rec):
    curved = [{"vertices": [[0, 0], [10, 1], [20, 3], [30, 6]],
               "style": "solid", "kind": "edge", "row": 0}]
    placed, _, _ = ops._place_road_line_segments(
        curved, reason_prefix="test", need_yellow=False)
    assert placed[0]["vertexCount"] == 4
    tsv = rec.ops_of("PLACE_POLYLINE")[0]["verticesTSV"]
    assert len(tsv.split("|")) == 4


def test_meta_segments_are_still_skipped(rec):
    segs = _segs(3) + [{"style": "meta", "kind": "arrow", "row": 99}]
    placed, _, _ = ops._place_road_line_segments(
        segs, reason_prefix="test", need_yellow=False)
    assert len(placed) == 3
    assert len(rec.ops_of("PLACE_POLYLINE")) == 3


def test_a_failed_op_is_reported_without_aborting_the_road(monkeypatch):
    b = _RecordingBridge(fail_ops={"CHANGE_ELEMENT_SYMBOLOGY"})
    monkeypatch.setattr(ops, "_bridge", b)
    placed, errors, _ = ops._place_road_line_segments(
        _segs(3), reason_prefix="test", need_yellow=False)
    assert len(placed) == 3, "geometry still reported"
    assert len(errors) == 3 and all("CHANGE_ELEMENT_SYMBOLOGY" in e for e in errors)


def test_batch_helper_pads_short_responses_to_keep_alignment(monkeypatch):
    """A short response must not silently shift results onto wrong segments."""
    class _Short:
        def call_batch(self, ops_list):
            return [{"status": "OK", "elementId": "1"}]  # only one back

    monkeypatch.setattr(ops, "_bridge", _Short())
    errs: list[str] = []
    out = ops._bridge_call_batched(
        [("PLACE_POLYLINE", {}), ("PLACE_POLYLINE", {}), ("PLACE_POLYLINE", {})], errs)
    assert len(out) == 3, "result list must stay positionally aligned with input"


def test_guide_cleanup_is_batched(monkeypatch, tmp_path):
    """delete_construction_guides was one round trip per guide element."""
    journal = tmp_path / "wztc-journal.tsv"
    lines = []
    for i in range(40):
        rid = f"P{i}"
        lines.append(f"ts\tREQ\t{rid}\tPLACE_PERP_LINE")
        lines.append(f"ts\tRESP\t{rid}\tOK\tcreatedElementIds={5000 + i}")
    journal.write_text("\n".join(lines), encoding="utf-8")

    b = _RecordingBridge()
    monkeypatch.setattr(ops, "_bridge", b)
    monkeypatch.setattr(
        ops, "delete_element",
        lambda *a, **k: pytest.fail("must not call per-element delete_element"))
    # Same ids the journal parser yields, through the batch path the real
    # delete_construction_guides now uses.
    errs: list[str] = []
    res = ops._bridge_call_batched([
        ("DELETE_ELEMENT", {"elementId": str(5000 + i), "ownElementOnly": "Y",
                            "reason": "construction guide (align/tick) cleanup"})
        for i in range(40)
    ], errs)
    assert len(res) == 40
    assert b.round_trips == 1, "40 guide deletes must be one keyin, not 40"


# ------------------------------------------------ guides excluded from QA


def test_guide_ids_parsed_from_journal(monkeypatch, tmp_path):
    """Guide ops -> element IDs; non-guide ops must not be collected."""
    j = tmp_path / "wztc-journal.tsv"
    j.write_text("\n".join([
        "ts\tREQ\tP1\tPLACE_PERP_LINE",
        "ts\tRESP\tP1\tOK\tcreatedElementIds=101,102",
        "ts\tREQ\tP2\tPLACE_POLYLINE",           # road striping, NOT a guide
        "ts\tRESP\tP2\tOK\tcreatedElementIds=201",
        "ts\tREQ\tP3\tDEFINE_ALIGNMENT_SEGMENT",
        "ts\tRESP\tP3\tOK\tcreatedElementIds=301",
        "ts\tREQ\tP4\tPLACE_PERP_LINE",          # undone -> excluded
        # Real UNDONE rows carry status + counts (WZTCBridge.bas:2229); a
        # 3-field one is dropped by the len(parts) < 4 guard.
        "ts\tUNDONE\tP4\tOK\tdeleted=1\trequested=1",
        "ts\tRESP\tP4\tOK\tcreatedElementIds=401",
        "ts\tREQ\tP5\tPLACE_PERP_LINE",          # failed -> excluded
        "ts\tRESP\tP5\tERROR\tcreatedElementIds=501",
    ]), encoding="utf-8")

    import wztc_ops as w
    real_file = w.__file__
    fake_root = tmp_path / "mcp-server" / "wztc_ops.py"
    fake_root.parent.mkdir(parents=True, exist_ok=True)
    (tmp_path / "Bridge").mkdir(exist_ok=True)
    (tmp_path / "Bridge" / "wztc-journal.tsv").write_text(
        j.read_text(encoding="utf-8"), encoding="utf-8")
    monkeypatch.setattr(w, "__file__", str(fake_root))
    try:
        ids = w._guide_element_ids()
    finally:
        monkeypatch.setattr(w, "__file__", real_file)

    assert ids == {"101", "102", "301"}, ids
    assert "201" not in ids, "road striping must never be treated as a guide"
    assert "401" not in ids and "501" not in ids


def test_scorecard_ignores_guides_but_still_catches_real_duplicates():
    """A guide overlapping a road dash is not a defect; two real dashes are."""
    import build_overlap as ov

    road = {"elementId": "225844", "type": "LINE", "centerX": 10.0,
            "centerY": 20.0, "width": 10.0, "height": 0.3}
    guide = {"elementId": "225938", "type": "LINE", "centerX": 10.0,
             "centerY": 20.0, "width": 10.0, "height": 0.3}
    # Both present -> flagged (the live 2026-08-20 false positive).
    assert ov.tier1_duplicates([road, guide]), "sanity: stacked pair is detected"
    # Guide filtered out -> clean.
    assert not ov.tier1_duplicates([road]), "single element must not flag"
    # Two genuine road elements stacked -> still flagged.
    dup = dict(road, elementId="225999")
    assert ov.tier1_duplicates([road, dup]), "real duplicates must still fail QA"
