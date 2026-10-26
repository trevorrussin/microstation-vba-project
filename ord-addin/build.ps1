# Build the OpenRoads add-in (C# 5, system csc.exe).    powershell -File ord-addin\build.ps1 [-Loader]
#   impl   -> bin\impl\WZTCCifImpl.<yyyyMMddHHmmss>.dll   (new unique file every build, nothing is overwritten; the loader
#             inside ORD picks the newest one on the next call, no ORD restart, no registry change)
#   -Loader-> bin\WZTCCif.dll.new                         (only when the loader itself changes; ORD must be closed to swap it in)
param([switch]$Loader)
$ErrorActionPreference = "Stop"
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$ord = "C:\Program Files\Bentley\OpenRoads Designer 2023.00\OpenRoadsDesigner"; $cif = "$ord\Cif"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$refs = @("$cif\Bentley.CifNET.SDK.4.0.dll","$cif\Bentley.CifNET.SDK.Edit.4.0.dll","$cif\Bentley.CifNET.GeometryModel.SDK.4.0.dll",
          "$cif\Bentley.CifNET.TerrainModel.4.0.dll","$cif\Bentley.CifNET.LinearGeometry.4.0.dll","$cif\Bentley.CifNET.DocCenter.Mstn.4.0.dll",
          "$ord\Bentley.DgnPlatformNET.dll","$ord\ustation.dll",
          "$ord\Bentley.GeometryNET.dll","$ord\Bentley.GeometryNET.Structs.dll","$ord\Bentley.GeometryNET.Common.dll","$ord\Assemblies\Bentley.TerrainModelNET.dll","System.Web.Extensions.dll")
$refArgs = $refs | ForEach-Object { "/r:$_" }
New-Item -ItemType Directory -Force "$here\bin\impl" | Out-Null
if ($Loader) {
    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    # the file must be named WZTCCif.dll: csc takes the assembly name from /out, and the registry expects "WZTCCif"
    New-Item -ItemType Directory -Force "$here\bin\loader.$stamp" | Out-Null
    & $csc /nologo /target:library /platform:x64 /optimize "/out:$here\bin\loader.$stamp\WZTCCif.dll" "$here\src\WZTCCif.cs"
    "built loader: $here\bin\loader.$stamp\WZTCCif.dll  (swap in with ORD closed: copy over bin\WZTCCif.dll)"
} else {
    $stamp = Get-Date -Format "yyyyMMddHHmmss"
    $out = "$here\bin\impl\WZTCCifImpl.$stamp.dll"
    & $csc /nologo /target:library /platform:x64 /optimize "/out:$out" @refArgs "$here\src\CifImpl.cs" "$here\src\CifOps.cs" "$here\src\CifOpsRead.cs" "$here\src\CifOpsCreate.cs" "$here\src\CifOpsCV.cs" "$here\src\CifOpsNB.cs" "$here\src\CifOpsExact.cs" "$here\src\CifOpsCivil.cs"
    if ($LASTEXITCODE -ne 0) { throw "csc failed" }
    "built impl: $out"
}
