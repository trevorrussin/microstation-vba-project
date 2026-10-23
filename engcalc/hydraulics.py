"""Deterministic drainage calculations (US customary). The QA gate re-runs these to
reproduce any value the agent reports (CALC-01). Each function returns
{"value", "units", "method", "inputs"} so results are self-describing.

Sources: Rational method and Manning's equation as used in NYSDOT HDM Ch 8 §8.3 / §8.7
(look up the governing criterion with search_standards; cite it on the task card).
"""
from __future__ import annotations

import math


def _out(value, units, method, **inputs):
    return {"value": value, "units": units, "method": method, "inputs": inputs}


def rational_q_cfs(c: float, i_in_per_hr: float, area_ac: float) -> dict:
    """Peak discharge Q = C i A (cfs; i in in/hr, A in acres)."""
    if not 0 < c <= 1:
        raise ValueError("runoff coefficient C must be in (0, 1]")
    if i_in_per_hr <= 0 or area_ac <= 0:
        raise ValueError("i and A must be positive")
    return _out(c * i_in_per_hr * area_ac, "cfs", "rational Q=CiA", c=c, i_in_per_hr=i_in_per_hr, area_ac=area_ac)


def manning_full_pipe_q(diameter_in: float, slope_ft_per_ft: float, n: float) -> dict:
    """Full-flow capacity of a circular pipe: Q = (1.486/n) A R^(2/3) S^(1/2) (cfs)."""
    if diameter_in <= 0 or slope_ft_per_ft <= 0 or n <= 0:
        raise ValueError("diameter, slope and n must be positive")
    d = diameter_in / 12.0
    a = math.pi * d * d / 4.0
    r = d / 4.0
    return _out((1.486 / n) * a * r ** (2.0 / 3.0) * math.sqrt(slope_ft_per_ft), "cfs",
                "Manning full pipe", diameter_in=diameter_in, slope_ft_per_ft=slope_ft_per_ft, n=n)


def manning_full_pipe_v(diameter_in: float, slope_ft_per_ft: float, n: float) -> dict:
    """Full-flow velocity V = (1.486/n) R^(2/3) S^(1/2) (ft/s)."""
    d = diameter_in / 12.0
    r = d / 4.0
    return _out((1.486 / n) * r ** (2.0 / 3.0) * math.sqrt(slope_ft_per_ft), "ft/s",
                "Manning full pipe velocity", diameter_in=diameter_in, slope_ft_per_ft=slope_ft_per_ft, n=n)


def required_diameter_in(q_cfs: float, slope_ft_per_ft: float, n: float,
                         sizes_in=(12, 15, 18, 21, 24, 27, 30, 33, 36, 42, 48, 54, 60, 66, 72)) -> dict:
    """Smallest standard diameter whose full-flow capacity >= Q."""
    for s in sizes_in:
        if manning_full_pipe_q(s, slope_ft_per_ft, n)["value"] >= q_cfs:
            return _out(float(s), "in", "smallest standard size with full-flow Q >= design Q",
                        q_cfs=q_cfs, slope_ft_per_ft=slope_ft_per_ft, n=n)
    raise ValueError(f"Q={q_cfs} cfs exceeds largest listed size")


def pipe_cover_ft(rim_or_grade_el: float, pipe_invert_el: float, diameter_in: float, wall_in: float = 0.0) -> dict:
    """Cover over pipe crown = grade - (invert + D + wall) (ft)."""
    crown = pipe_invert_el + (diameter_in + wall_in) / 12.0
    return _out(rim_or_grade_el - crown, "ft", "cover = grade - crown",
                rim_or_grade_el=rim_or_grade_el, pipe_invert_el=pipe_invert_el, diameter_in=diameter_in, wall_in=wall_in)
