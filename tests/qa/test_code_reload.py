"""reload_code: edited mcp-server modules are reloaded in the running server without a restart, live state survives, a syntax error reloads nothing."""
import importlib
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

import code_reload  # noqa: E402
import workflow_contract  # noqa: E402,F401  (only modules already imported are reloaded)
import wztc_ops  # noqa: E402


def test_reload_keeps_session_state_and_bridge():
    sentinel = object()
    session = wztc_ops._PLAN_SESSION
    wztc_ops._LAST_PLACED_ROAD = {"marker": 1}
    out = code_reload.reload_code(sentinel)
    assert out["status"] == "OK", out
    assert "wztc_ops" in out["reloaded"] and "workflow_contract" in out["reloaded"]
    assert wztc_ops._PLAN_SESSION is session
    assert type(session) is wztc_ops.PlanSession                      # live object got the new class
    assert wztc_ops._LAST_PLACED_ROAD == {"marker": 1}
    assert wztc_ops._bridge is sentinel                                # bridge re-linked
    wztc_ops.set_bridge(None)


def test_reload_picks_up_an_edit_and_refuses_a_syntax_error(tmp_path):
    f = tmp_path / "hot_probe_mod.py"
    f.write_text("VALUE = 1\n", encoding="utf-8")
    sys.path.insert(0, str(tmp_path))
    try:
        mod = importlib.import_module("hot_probe_mod")
        assert mod.VALUE == 1
        f.write_text("VALUE = 2\n", encoding="utf-8")
        os.utime(f, (1_900_000_000, 1_900_000_000))                    # same size within one second would reuse the cached bytecode
        out =code_reload.reload_code(None, extra=("hot_probe_mod",))
        assert out["status"] == "OK" and mod.VALUE == 2
        f.write_text("VALUE = = 3\n", encoding="utf-8")
        out = code_reload.reload_code(None, extra=("hot_probe_mod",))
        assert out["status"] == "ERROR" and out["reloaded"] == [] and "hot_probe_mod" in out["syntax_errors"]
        assert mod.VALUE == 2                                          # nothing was reloaded
    finally:
        sys.path.remove(str(tmp_path))
        sys.modules.pop("hot_probe_mod", None)


def test_bridge_client_reload_keeps_live_instances_and_state():
    import bridge_client
    live, chat, tls, counter, err = bridge_client.bridge, bridge_client.chat_bridge, bridge_client._tls, bridge_client._req_counter, bridge_client.BridgeError
    out = code_reload.reload_code(None)
    assert out["status"] == "OK" and "bridge_client" in out["reloaded"], out
    assert bridge_client.bridge is live and bridge_client.chat_bridge is chat      # server.py / wztc_ops hold these objects
    assert type(live) is bridge_client.Bridge and type(chat) is bridge_client.Bridge  # ...which now run the new methods
    assert bridge_client._tls is tls and bridge_client._req_counter is counter     # COM handles and request ids are not reset
    assert bridge_client.BridgeError is err                                        # `except BridgeError` in other modules still matches
