"""Pydantic schema for Data/sheet-specs/<sheet>.json.

Structural shape only — catches missing required sections / mistyped
top-level fields early. Domain invariants (skip-line arithmetic, zone
cross-refs, resolution) stay in scripts/validate_sheet_spec.py; this is
the typed gate in front of them.

Uses extra='allow' throughout so authored fields can evolve without a
schema chase on every sheet family. Nested table row shapes stay
dict[str, Any] for the same reason.
"""
from __future__ import annotations

from typing import Any, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator


class SheetMeta(BaseModel):
    model_config = ConfigDict(extra="allow")
    number: str
    title: str
    kind: Optional[Literal["plan", "referenceLibrary"]] = None


class CorridorZone(BaseModel):
    model_config = ConfigDict(extra="allow")
    id: str
    order: int
    kind: str


class Corridor(BaseModel):
    model_config = ConfigDict(extra="allow")
    zones: list[CorridorZone] = Field(min_length=1)


class OrderTableRow(BaseModel):
    model_config = ConfigDict(extra="allow")
    rowNum: int
    type: Literal["Sign", "Non-Sign"]


class OrderAlignment(BaseModel):
    model_config = ConfigDict(extra="allow")
    alignIdx: int
    rows: list[OrderTableRow] = Field(min_length=1)
    overlayZones: list[dict[str, Any]] = Field(default_factory=list)


class OrderTable(BaseModel):
    model_config = ConfigDict(extra="allow")
    alignments: list[OrderAlignment] = Field(min_length=1)


class SignItem(BaseModel):
    model_config = ConfigDict(extra="allow")
    signCode: str


class Signs(BaseModel):
    model_config = ConfigDict(extra="allow")
    items: list[SignItem] = Field(min_length=1)


class RuleItem(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)
    id: str
    severity: Literal["error", "warning"]
    source: str
    # JSON key is "assert"; alias keeps the Python attr legal.
    assert_: str = Field(alias="assert")
    commonFailure: str


class SheetSpec(BaseModel):
    """Root model. Plan sheets require corridor/orderTable/signs/etc.;
    reference-library sheets only require tables + tableRoles + sheet."""
    model_config = ConfigDict(extra="allow")

    schemaVersion: str
    sheet: SheetMeta
    tableRoles: dict[str, Any]
    tables: dict[str, Any]

    applicability: Optional[dict[str, Any]] = None
    inputs: Optional[list[dict[str, Any]]] = None
    corridor: Optional[Corridor] = None
    orderTable: Optional[OrderTable] = None
    signs: Optional[Signs] = None
    symbols: Optional[dict[str, Any]] = None
    annotations: Optional[dict[str, Any]] = None
    annotationStyle: Optional[dict[str, Any]] = None
    rules: Optional[list[RuleItem]] = None
    notes: Optional[Any] = None
    geometry: Optional[dict[str, Any]] = None
    details: Optional[Any] = None
    legend: Optional[dict[str, Any]] = None
    knownExcerpts: Optional[Any] = None
    knownCodeDeviations: Optional[Any] = None
    knownAnomalies: Optional[Any] = None
    openQuestions: Optional[Any] = None

    @model_validator(mode="after")
    def _plan_or_library_sections(self) -> "SheetSpec":
        kind = self.sheet.kind or "plan"
        if kind == "referenceLibrary":
            if not self.tables:
                raise ValueError("referenceLibrary sheet requires non-empty tables")
            return self
        missing = [
            name for name in (
                "applicability", "corridor", "orderTable", "signs",
                "symbols", "annotations", "rules",
            )
            if getattr(self, name) is None
        ]
        if missing:
            raise ValueError(
                f"plan sheet requires sections: {', '.join(missing)}"
            )
        if not self.tables:
            raise ValueError("plan sheet requires non-empty tables")
        return self


def validate_sheet_dict(raw: dict) -> SheetSpec:
    """Raise pydantic.ValidationError if the dict is not a valid sheet spec."""
    return SheetSpec.model_validate(raw)


def format_validation_error(exc: ValidationError) -> str:
    parts = []
    for err in exc.errors():
        loc = ".".join(str(x) for x in err.get("loc", ()))
        parts.append(f"{loc}: {err.get('msg')}")
    return "; ".join(parts) if parts else str(exc)
