# Register WZTC.CifBridge for the CURRENT USER only (HKCU, no admin). Run after every rebuild only if the path changes.
#   powershell -File ord-addin\register.ps1            register
#   powershell -File ord-addin\register.ps1 -Remove    unregister
param([switch]$Remove)
$guid = "{8D3A5C1E-6F42-4B7A-9E10-2C7B5D9A41F3}"
$progId = "WZTC.CifBridge"
$dll = (Resolve-Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bin\WZTCCif.dll")).Path
$cls = "HKCU:\Software\Classes"
if ($Remove) {
    Remove-Item "$cls\$progId" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$cls\CLSID\$guid" -Recurse -Force -ErrorAction SilentlyContinue
    "unregistered $progId"
    return
}
New-Item "$cls\$progId" -Force | Out-Null
Set-ItemProperty "$cls\$progId" -Name "(default)" -Value $progId
New-Item "$cls\$progId\CLSID" -Force | Out-Null
Set-ItemProperty "$cls\$progId\CLSID" -Name "(default)" -Value $guid
New-Item "$cls\CLSID\$guid" -Force | Out-Null
Set-ItemProperty "$cls\CLSID\$guid" -Name "(default)" -Value $progId
$in = "$cls\CLSID\$guid\InprocServer32"
New-Item $in -Force | Out-Null
Set-ItemProperty $in -Name "(default)" -Value "mscoree.dll"
Set-ItemProperty $in -Name "ThreadingModel" -Value "Both"
Set-ItemProperty $in -Name "Class" -Value "WZTC.CifBridge"
Set-ItemProperty $in -Name "Assembly" -Value "WZTCCif, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty $in -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty $in -Name "CodeBase" -Value ("file:///" + ($dll -replace "\\", "/"))
New-Item "$cls\CLSID\$guid\ProgId" -Force | Out-Null
Set-ItemProperty "$cls\CLSID\$guid\ProgId" -Name "(default)" -Value $progId
"registered $progId -> $dll"
