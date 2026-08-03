import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "mcp-server"))
import sheet_spec as ss  # noqa: E402

sp = ss.load("619-311")
r = ss.resolve(sp, 45, 12, "12 ft", "URBAN", "LANE CLOSURE", "WORKERS ON FOOT")
print("resolved :", {k: v for k, v in r.items() if k != "legend"})
print("legend   :", r["legend"])

p = ss.order_table_rows(sp, r)
print("\nnonSignRows:")
for x in p["nonSignRows"]:
    print("   ", x)
print("signRows:")
for x in p["signRows"]:
    print("   ", x)
print("overlays:")
for x in p["overlays"]:
    print("   ", x)

print("\nstation walk:")
for w in ss.station_walk(sp, r):
    print(f"  {w['alignName']:<10} {str(w['item']):<24} "
          f"len={w['lengthFt']:>7g}  sta={w['stationFt']:>7g}  {w.get('note', '')}")

print("\n--- guard rails ---")
for args, why in (
    ((sp, 65, 12, "12 ft", "URBAN"), "speed outside the sheet's table range"),
    ((sp, 45, 12, "wide", "URBAN"), "uninterpretable shoulder width"),
    ((sp, 45, 13, "12 ft", "URBAN"), "lane width the sheet doesn't print"),
):
    try:
        ss.resolve(*args)
        print(f"  NO ERROR RAISED for {why} <-- BAD")
    except ss.SpecError as e:
        print(f"  raises on {why}:\n      {e}")

print("\n--- RURAL variant (legend changes the SignLibrary key) ---")
rr = ss.resolve(sp, 45, 12, "12 ft", "RURAL")
print("  legend   :", rr["legend"])
print("  signRows :", ss.order_table_rows(sp, rr)["signRows"])

print("\n--- shoulder band collapse (the ComputeSpacing fabrication) ---")
for key in ("<= 4 ft", "5-7 ft", "8 ft", "10 ft", "12 ft"):
    band = ss.shoulder_band(sp, key)
    ft = ss.resolve(sp, 45, 12, key, "URBAN")["shoulderTaper"]["ft"]
    print(f"  app {key:<9} -> sheet band {band:<9} -> shoulder taper {ft} ft")
print("  (ComputeSpacing returns 160 ft for 45 mph / 12 ft shoulder; "
      "Table 311-02 says 120)")
