"""Declarative parameter spec for the road striping catalog.

Mirrors the pattern that already works for named 619 sheets: those carry a
machine-readable `inputs[]` array (id / label / allowed / default), so
`get_required_designer_inputs` can ask exactly the missing questions instead
of the agent inferring them from prose. The striping tools had no equivalent,
so "build me a curved highway" and "S-curve, 2000 ft, 2 bends, 4 lanes" both
went through the same unstructured guesswork -- the vague user got
interrogated, and the specific user risked being re-asked what they just said.

Three classes of parameter:
  REQUIRED    -- only the engineer knows it; must ask (or propose + confirm)
  DEFAULTABLE -- safe standard value, but it must be ANNOUNCED, never silent
  DERIVED     -- computed from another answer; never ask

`path` is modelled as one required input satisfied several ways (synthesize,
reuse this session's road, pick an element, click points), so a vague request
resolves to a proposal rather than a click session.
"""
from __future__ import annotations

from typing import Any, Optional

REQUIRED = "required"
DEFAULTABLE = "defaultable"
DERIVED = "derived"


def _lanes(label: str, allowed: list[int]) -> dict:
    return {"id": "lanes", "label": label, "kind": REQUIRED, "type": "integer",
            "allowed": allowed}


_PATH_INPUT = {
    "id": "path",
    "label": "Roadway path (alignment)",
    "kind": REQUIRED,
    "type": "path",
    "note": ("Satisfied by vertices=, by endpoints x1..y2, or by "
             "propose_road_path when the engineer only described a shape."),
    "options": [
        {"label": "Let me propose one", "value": "synthesize",
         "description": "propose_road_path(length_ft, kind/bends) — for "
                        "'a curved highway' with no drawn geometry."},
        {"label": "The road I just placed", "value": "last_placed",
         "description": "Reuse this session's striping vertices (no click)."},
        {"label": "I'll click the roadway", "value": "element",
         "description": "One element pick → get_element_vertices."},
        {"label": "I'll click the endpoints", "value": "points",
         "description": "Two point picks (straight runs only)."},
    ],
}

_COMMON_DEFAULTS = [
    {"id": "lane_width_ft", "label": "Lane width (ft)", "kind": DEFAULTABLE,
     "type": "number", "default": 12.0, "allowed": [10, 11, 12]},
    {"id": "shoulder_width_ft", "label": "Paved shoulder width (ft)",
     "kind": DEFAULTABLE, "type": "number", "default": 0.0},
    {"id": "dash_ft", "label": "Lane-line dash (ft)", "kind": DEFAULTABLE,
     "type": "number", "default": 10.0},
    {"id": "gap_ft", "label": "Lane-line gap (ft)", "kind": DEFAULTABLE,
     "type": "number", "default": 30.0},
    {"id": "side", "label": "Which side of the path is travel", "kind": DEFAULTABLE,
     "type": "enum", "default": "right", "allowed": ["right", "left"],
     "note": ("The path is the FIRST TRAVEL OUTER EDGE, not the centreline. "
              "Getting this backwards is what put cones through the middle of "
              "the road — confirm it rather than defaulting silently.")},
]

_YELLOW_GAP = {"id": "yellow_gap_ft", "label": "Gap between double yellows (ft)",
               "kind": DEFAULTABLE, "type": "number", "default": 2.0}


ROAD_TOOL_INPUTS: dict[str, dict[str, Any]] = {
    "place_lane_highway": {
        "label": "One-way / single carriageway",
        "inputs": [_lanes("Number of travel lanes", [1, 2, 3, 4, 5, 6]),
                   _PATH_INPUT] + _COMMON_DEFAULTS,
    },
    "place_two_way_highway": {
        "label": "Undivided two-way, double solid yellow",
        "inputs": [
            {"id": "lanes", "label": "Total lanes (both directions, even)",
             "kind": REQUIRED, "type": "integer", "allowed": [2, 4, 6, 8]},
            _PATH_INPUT, _YELLOW_GAP] + _COMMON_DEFAULTS,
    },
    "place_divided_highway": {
        "label": "Divided / multilane with physical median",
        "inputs": [
            {"id": "lanes_per_direction", "label": "Lanes per direction",
             "kind": REQUIRED, "type": "integer", "allowed": [1, 2, 3, 4]},
            {"id": "median_width_ft", "label": "Median width (ft)",
             "kind": REQUIRED, "type": "number",
             "note": "Required — a divided road is not defined without it."},
            _PATH_INPUT] + _COMMON_DEFAULTS,
    },
    "place_twlt_highway": {
        "label": "Multilane undivided with centre two-way left-turn lane",
        "inputs": [
            {"id": "lanes_per_direction", "label": "Lanes per direction",
             "kind": REQUIRED, "type": "integer", "allowed": [1, 2, 3]},
            {"id": "twlt_width_ft", "label": "TWLT width (ft)", "kind": DEFAULTABLE,
             "type": "number", "default": 12.0},
            _PATH_INPUT] + _COMMON_DEFAULTS,
    },
    "place_ramp_gore": {
        "label": "Freeway ramp gore / diverge",
        "inputs": [
            _PATH_INPUT,
            {"id": "ramp_angle_deg", "label": "Ramp departure angle (deg)",
             "kind": REQUIRED, "type": "number"},
            {"id": "gore_station_ft", "label": "Gore station along mainline (ft)",
             "kind": REQUIRED, "type": "number"},
            {"id": "ramp_length_ft", "label": "Ramp length (ft)",
             "kind": REQUIRED, "type": "number"},
            {"id": "mainline_lanes", "label": "Mainline lanes", "kind": DEFAULTABLE,
             "type": "integer", "default": 2},
            {"id": "ramp_lanes", "label": "Ramp lanes", "kind": DEFAULTABLE,
             "type": "integer", "default": 1},
            {"id": "gore_mark_ft", "label": "Gore chevron mark (ft)",
             "kind": DEFAULTABLE, "type": "number", "default": 40.0},
        ] + _COMMON_DEFAULTS,
    },
    "place_orthogonal_intersection": {
        "label": "Orthogonal + or T intersection",
        "inputs": [
            {"id": "junction_point", "label": "Junction point (x, y)",
             "kind": REQUIRED, "type": "point",
             "note": "Ask for a point pick — never invent site coordinates."},
            {"id": "junction", "label": "Plus or tee", "kind": REQUIRED,
             "type": "enum", "allowed": ["plus", "tee"]},
            {"id": "tee_side", "label": "Which side the stub leaves",
             "kind": REQUIRED, "type": "enum", "allowed": ["right", "left"],
             "requiredWhen": {"junction": "tee"}},
            {"id": "primary_road_type", "label": "Primary arm road type",
             "kind": REQUIRED, "type": "enum",
             "allowed": ["one_way", "two_way", "divided", "twlt"]},
            {"id": "secondary_road_type", "label": "Secondary arm road type",
             "kind": REQUIRED, "type": "enum",
             "allowed": ["one_way", "two_way", "divided", "twlt"]},
            {"id": "primary_length_ft", "label": "Primary arm length (ft)",
             "kind": REQUIRED, "type": "number"},
            {"id": "secondary_stub_ft", "label": "Secondary stub length (ft)",
             "kind": REQUIRED, "type": "number"},
            {"id": "primary_median_width_ft", "label": "Primary median width (ft)",
             "kind": REQUIRED, "type": "number",
             "requiredWhen": {"primary_road_type": "divided"}},
            {"id": "secondary_median_width_ft", "label": "Secondary median width (ft)",
             "kind": REQUIRED, "type": "number",
             "requiredWhen": {"secondary_road_type": "divided"}},
            # Not DEFAULTABLE: its default is "same as lanes_in", which is
            # another answer rather than a constant. Never ask for it — set it
            # only when the engineer says a lane drops.
            {"id": "primary_lanes_out", "label": "Primary lanes leaving the box",
             "kind": DERIVED, "type": "integer",
             "derivedFrom": ("lanes_in (all through) unless the engineer says a "
                             "lane drops; a drop is what turns shared arrows "
                             "into dedicated SAL/SAR + SLONLY")},
            {"id": "secondary_lanes_out", "label": "Secondary lanes leaving the box",
             "kind": DERIVED, "type": "integer",
             "derivedFrom": "lanes_in unless the engineer says a lane drops"},
            {"id": "has_turning_lanes", "label": "Dotted yellow through the box",
             "kind": DERIVED, "type": "boolean",
             "derivedFrom": "true when a TWLT arm is present or dedicated > 0"},
            {"id": "primary_bearing_deg", "label": "Primary arm bearing (deg)",
             "kind": DEFAULTABLE, "type": "number", "default": 0.0},
            {"id": "crosswalks", "label": "Crosswalks", "kind": DEFAULTABLE,
             "type": "boolean", "default": True},
            {"id": "stop_bars", "label": "Stop bars", "kind": DEFAULTABLE,
             "type": "boolean", "default": True},
            {"id": "turn_arrows", "label": "Turn arrows", "kind": DEFAULTABLE,
             "type": "boolean", "default": True},
            {"id": "lane_width_ft", "label": "Lane width (ft)", "kind": DEFAULTABLE,
             "type": "number", "default": 12.0},
        ],
    },
}

# Engineer wording -> tool. Keeps "build me a four lane road" from having to
# name a Python function.
_TOOL_ALIASES = {
    "one_way": "place_lane_highway", "oneway": "place_lane_highway",
    "freeway": "place_lane_highway", "lane_highway": "place_lane_highway",
    "two_way": "place_two_way_highway", "twoway": "place_two_way_highway",
    "undivided": "place_two_way_highway", "highway": "place_two_way_highway",
    "road": "place_two_way_highway",
    "divided": "place_divided_highway", "median": "place_divided_highway",
    "twlt": "place_twlt_highway", "turn_lane": "place_twlt_highway",
    "intersection": "place_orthogonal_intersection",
    "junction": "place_orthogonal_intersection",
    "ramp": "place_ramp_gore", "gore": "place_ramp_gore",
}


def resolve_tool(name: str) -> str:
    """Map a tool name or an engineer's word to a catalog tool."""
    n = (name or "").strip().lower().replace("-", "_").replace(" ", "_")
    if n in ROAD_TOOL_INPUTS:
        return n
    return _TOOL_ALIASES.get(n, "")


def _is_answered(known: dict, key: str) -> bool:
    if key == "path":
        return bool(known.get("path") or known.get("vertices")
                    or known.get("path_vertices")
                    or (known.get("x2") is not None and known.get("y2") is not None))
    if key == "junction_point":
        return known.get("junction_x") is not None and known.get("junction_y") is not None
    v = known.get(key)
    return v is not None and v != ""


def _required_now(inp: dict, known: dict) -> bool:
    """Conditional requirements only bite once their trigger is answered."""
    if inp.get("kind") != REQUIRED:
        return False
    cond = inp.get("requiredWhen")
    if not cond:
        return True
    return all(str(known.get(k, "")).lower() == str(v).lower() for k, v in cond.items())


def get_required_road_inputs(tool: str, known: Optional[dict] = None) -> dict:
    """Return only the gaps, plus every default that will be applied.

    Ask nothing the engineer already said; announce everything assumed.
    """
    known = dict(known or {})
    resolved = resolve_tool(tool)
    if not resolved:
        return {"status": "ERROR", "found": False, "tool": tool,
                "note": (f"unknown road tool {tool!r}; expected one of "
                         f"{sorted(ROAD_TOOL_INPUTS)}")}
    spec = ROAD_TOOL_INPUTS[resolved]
    missing: list[dict] = []
    defaults: list[dict] = []
    derived: list[dict] = []
    answered: list[str] = []
    for inp in spec["inputs"]:
        iid = inp["id"]
        if _is_answered(known, iid):
            answered.append(iid)
            continue
        kind = inp.get("kind")
        if kind == DERIVED:
            derived.append({"id": iid, "label": inp["label"],
                            "derivedFrom": inp.get("derivedFrom", "")})
        elif kind == DEFAULTABLE:
            defaults.append({"id": iid, "label": inp["label"],
                             "value": inp.get("default"),
                             "note": inp.get("note", "")})
        elif _required_now(inp, known):
            row = {"id": iid, "label": inp["label"], "type": inp.get("type")}
            for k in ("allowed", "options", "note"):
                if inp.get(k):
                    row[k] = inp[k]
            missing.append(row)
    return {
        "status": "OK",
        "tool": resolved,
        "toolLabel": spec["label"],
        "missing": missing,
        "answered": answered,
        "assumedDefaults": defaults,
        "derived": derived,
        "ready": not missing,
        "note": (
            "Ask one ask_user_choice per missing row, using its allowed/options "
            "verbatim. State assumedDefaults in your reply — never apply them "
            "silently. When only `path` is missing and the engineer described a "
            "shape, call propose_road_path and confirm instead of asking."
        ),
    }


def announce_defaults(assumed: list[dict]) -> str:
    """One readable sentence for the agent's reply."""
    if not assumed:
        return ""
    parts = []
    for d in assumed:
        v = d.get("value")
        if isinstance(v, bool):
            parts.append(f"{d['label'].lower()} {'on' if v else 'off'}")
        else:
            parts.append(f"{d['label'].lower()} {v}")
    return "Assumed: " + ", ".join(parts) + "."
