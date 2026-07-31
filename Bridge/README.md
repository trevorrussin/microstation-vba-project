# Bridge

File-based transport for `Modules/WZTCBridge.bas`. An external process (or, for
manual M1 testing, a human with a text editor) writes `request.tsv`, triggers
`VBA RUN [ProjectName]WZTCBridge.RunRequest` in MicroStation, and reads
`response.tsv` for the result. Every executed op is also appended to
`wztc-journal.tsv`.

These three `.tsv` files are runtime artifacts (git-ignored) — this README is
what keeps the folder itself in the repo.

## Protocol

One op per line, tab-separated:

```
<reqId>\t<OP_TYPE>\tkey1=val1\tkey2=val2...
```

Response mirrors it, one line per request, in the same order:

```
<reqId>\t<OK|ERROR>\tkey1=val1...
```

## M1 op: PLACE_CELL

```
0001	PLACE_CELL	cellName=TWZAP_P	ptX=1000	ptY=1000	ptZ=0	angleDeg=0
```

`cellName` must be a valid entry from `CellPlacer.GetCellCatalogue()` (e.g.
`TWZAP_P` — Arrow Panel). `ptZ` and `angleDeg` are optional, default `0`.

Success response:

```
0001	OK	elementId=88213	note=placed TWZAP_P at 1000,1000
```

## Manual test (no Python yet)

1. Open a design file with the WZTC VBA project loaded.
2. Create `request.tsv` in this folder with the line above.
3. In MicroStation's Key-in bar, type:
   `VBA RUN [ProjectName]WZTCBridge.RunRequest`
   (replace `[ProjectName]` with the actual project name shown in the VBA IDE)
4. Confirm a `TWZAP_P` cell appeared at design coordinates (1000, 1000).
5. Check `response.tsv` — should show `OK` and an `elementId`.
6. Check `wztc-journal.tsv` — should show the request and response lines appended.

This proves the VBA-side half of the bridge works before wiring up an
external Python client to trigger it over COM.
