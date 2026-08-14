"""CLEAR_PLAN must recover createdElementIds after journal rotation."""
from __future__ import annotations

import wztc_ops as ops


def test_harvest_binds_resp_to_latest_req_not_global_last_wins():
    # Same reqId P1 reused: first PLACE_TEXT leftover, then later HANDOFF.
    text = (
        "t\tREQ\tP1\tPLACE_TEXT_LABEL\talignIdx=1\n"
        "t\tRESP\tP1\tOK\tcreatedElementIds=164238\n"
        "t\tREQ\tP1\tHANDOFF\n"
        "t\tRESP\tP1\tOK\n"
        "t\tREQ\tP2\tPLACE_DIMENSION\talignIdx=1\n"
        "t\tRESP\tP2\tOK\tcreatedElementIds=183625\n"
    )
    ids = ops.harvest_journal_create_ids(text, keep_alignments=True)
    assert ids == {"164238", "183625"}


def test_harvest_skips_align_ops_when_keeping_corridor():
    text = (
        "t\tREQ\tP1\tDEFINE_ALIGNMENT_SEGMENT\n"
        "t\tRESP\tP1\tOK\tcreatedElementIds=100\n"
        "t\tREQ\tP3\tPLACE_TEXT_LABEL\n"
        "t\tRESP\tP3\tOK\tcreatedElementIds=300\n"
    )
    ids = ops.harvest_journal_create_ids(text, keep_alignments=True)
    assert ids == {"300"}
    ids2 = ops.harvest_journal_create_ids(text, keep_alignments=False)
    assert ids2 == {"100", "300"}


def test_harvest_respects_align_idx_scope():
    text = (
        "t\tREQ\tP1\tPLACE_SIGN\talignIdx=1\n"
        "t\tRESP\tP1\tOK\tcreatedElementIds=11\n"
        "t\tREQ\tP2\tPLACE_SIGN\talignIdx=2\n"
        "t\tRESP\tP2\tOK\tcreatedElementIds=22\n"
    )
    assert ops.harvest_journal_create_ids(text, align_idx=2) == {"22"}
    assert ops.harvest_journal_create_ids(text, align_idx=1) == {"11"}
    assert ops.harvest_journal_create_ids(text) == {"11", "22"}
