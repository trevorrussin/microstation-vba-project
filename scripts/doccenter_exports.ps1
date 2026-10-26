# Read-only: list PE export names from ORD DocCenter (no LoadLibrary, no UAC).
# Usage: powershell -File scripts/doccenter_exports.ps1 [-Filter BuildProfile]
param([string]$Filter = "BuildProfile|BuildCifProfile|CreateNamedBoundary|CreateNewDrawing")
$ErrorActionPreference = "Stop"
$dll = "C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner\Cif\Bentley.Cif.DocCenter.4.0.dll"
if (-not (Test-Path $dll)) { throw "DocCenter not found: $dll" }

function Get-PeExports([string]$Path) {
  $fs = [System.IO.File]::OpenRead($Path)
  $br = New-Object System.IO.BinaryReader($fs)
  try {
    $fs.Seek(0x3C, 'Begin') | Out-Null
    $peOff = $br.ReadInt32()
    $fs.Seek($peOff + 4, 'Begin') | Out-Null
    $br.ReadUInt16() | Out-Null
    $numSections = $br.ReadUInt16()
    $br.ReadUInt32() | Out-Null; $br.ReadUInt32() | Out-Null; $br.ReadUInt32() | Out-Null
    $optSize = $br.ReadUInt16(); $br.ReadUInt16() | Out-Null
    $optStart = $peOff + 24
    $fs.Seek($optStart, 'Begin') | Out-Null
    $magic = $br.ReadUInt16()
    $ddOff = $optStart + $(if ($magic -eq 0x20b) { 112 } else { 96 })
    $fs.Seek($ddOff, 'Begin') | Out-Null
    $exportRva = $br.ReadUInt32()
    if ($exportRva -eq 0) { return @() }
    $secOff = $optStart + $optSize
    $sections = @()
    for ($i = 0; $i -lt $numSections; $i++) {
      $fs.Seek($secOff + $i * 40 + 8, 'Begin') | Out-Null
      $vsize = $br.ReadUInt32(); $vaddr = $br.ReadUInt32()
      $rsize = $br.ReadUInt32(); $raw = $br.ReadUInt32()
      $sections += [pscustomobject]@{ VAddr = $vaddr; VSize = $vsize; Raw = $raw; RSize = $rsize }
    }
    $script:sections = $sections
    function RvaToOff([uint32]$rva) {
      foreach ($s in $script:sections) {
        $span = [Math]::Max($s.VSize, $s.RSize)
        if ($rva -ge $s.VAddr -and $rva -lt ($s.VAddr + $span)) {
          return [int64]($rva - $s.VAddr + $s.Raw)
        }
      }
      throw "RVA $rva not in sections"
    }
    $expOff = RvaToOff $exportRva
    $fs.Seek($expOff, 'Begin') | Out-Null
    $br.ReadUInt32() | Out-Null; $br.ReadUInt32() | Out-Null
    $br.ReadUInt16() | Out-Null; $br.ReadUInt16() | Out-Null
    $br.ReadUInt32() | Out-Null
    $ordinalBase = $br.ReadUInt32()
    $br.ReadUInt32() | Out-Null
    $nNames = $br.ReadUInt32()
    $br.ReadUInt32() | Out-Null
    $addrOfNames = $br.ReadUInt32()
    $addrOfOrds = $br.ReadUInt32()
    $names = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $nNames; $i++) {
      $fs.Seek((RvaToOff $addrOfNames) + 4 * $i, 'Begin') | Out-Null
      $nrva = $br.ReadUInt32()
      $fs.Seek((RvaToOff $addrOfOrds) + 2 * $i, 'Begin') | Out-Null
      $ord = $br.ReadUInt16()
      $fs.Seek((RvaToOff $nrva), 'Begin') | Out-Null
      $chars = New-Object System.Collections.Generic.List[byte]
      while (($b = $br.ReadByte()) -ne 0) { $chars.Add($b) }
      $nm = [System.Text.Encoding]::ASCII.GetString($chars.ToArray())
      $names.Add([pscustomobject]@{ Ordinal = ($ordinalBase + $ord); Name = $nm })
    }
    return $names
  } finally { $br.Close(); $fs.Close() }
}

$all = Get-PeExports $dll
$hit = $all | Where-Object { $_.Name -match $Filter }
"dll=$dll"
"total_exports=$($all.Count) matched=$($hit.Count)"
$hit | ForEach-Object { "$($_.Ordinal)`t$($_.Name)" }
$outDir = Join-Path (Split-Path $PSScriptRoot) "outputs\tooling"
New-Item -ItemType Directory -Force $outDir | Out-Null
$out = Join-Path $outDir "doccenter-create-exports.txt"
$hit | ForEach-Object { "$($_.Ordinal)`t$($_.Name)" } | Set-Content $out -Encoding utf8
"wrote $out"
