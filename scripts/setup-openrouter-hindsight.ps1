# Wire Hindsight (local daemon on :8888) to OpenRouter.
# Does NOT print the API key. Does NOT configure Cursor chat (UI-only).
#
# Usage (pick one):
#   1) $env:OPENROUTER_API_KEY = "sk-or-v1-..."
#      powershell -ExecutionPolicy Bypass -File scripts\setup-openrouter-hindsight.ps1
#   2) powershell -ExecutionPolicy Bypass -File scripts\setup-openrouter-hindsight.ps1
#      (script prompts; paste key — nothing is echoed)
#
# Optional:
#   -Model "qwen/qwen3-14b"   (OpenRouter model id)
#   -SkipRestart              (update config only)

param(
    [string]$Model = "qwen/qwen3-14b",
    [switch]$SkipRestart
)

$ErrorActionPreference = "Stop"

function Get-OpenRouterKey {
    $fromEnv = [Environment]::GetEnvironmentVariable("OPENROUTER_API_KEY", "Process")
    if (-not $fromEnv) {
        $fromEnv = [Environment]::GetEnvironmentVariable("OPENROUTER_API_KEY", "User")
    }
    if ($fromEnv) { return $fromEnv }

    Write-Host "Paste OpenRouter API key (input hidden), then Enter:"
    $secure = Read-Host -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

$key = Get-OpenRouterKey
if (-not $key -or $key -notmatch '^sk-or-') {
    throw "Expected an OpenRouter key starting with sk-or-. Aborting."
}

if (-not (Get-Command hindsight-embed -ErrorAction SilentlyContinue)) {
    throw "hindsight-embed not found on PATH."
}

Write-Host "Setting Hindsight profile env (provider=openrouter, model=$Model)..."
hindsight-embed profile set-env default HINDSIGHT_API_LLM_PROVIDER openrouter
hindsight-embed profile set-env default HINDSIGHT_API_LLM_API_KEY $key
hindsight-embed profile set-env default HINDSIGHT_API_LLM_MODEL $Model
# Clear Ollama-only base URL if present so OpenRouter default applies
hindsight-embed profile remove-env default HINDSIGHT_API_LLM_BASE_URL 2>$null

$cursorJson = Join-Path $env:USERPROFILE ".hindsight\cursor.json"
if (Test-Path $cursorJson) {
    Write-Host "Updating $cursorJson ..."
    $cfg = Get-Content $cursorJson -Raw | ConvertFrom-Json
    $cfg.llmProvider = "openrouter"
    $cfg.llmModel = $Model
    # Do not store the API key in this JSON (profile env holds it)
    $cfg | ConvertTo-Json -Depth 8 | Set-Content $cursorJson -Encoding utf8
} else {
    Write-Warning "Missing $cursorJson — profile env still set."
}

# Persist key for future shells (User scope). Never write into the git repo.
[Environment]::SetEnvironmentVariable("OPENROUTER_API_KEY", $key, "User")
Write-Host "Saved OPENROUTER_API_KEY to User environment (new terminals pick it up)."

if (-not $SkipRestart) {
    Write-Host "Restarting Hindsight daemon..."
    hindsight-embed daemon stop 2>$null
    Start-Sleep -Seconds 1
    hindsight-embed daemon start
    Start-Sleep -Seconds 2
    hindsight-embed daemon status
}

Write-Host ""
Write-Host "Hindsight is on OpenRouter model: $Model"
Write-Host "Cursor general chat is separate — see chat instructions (Settings → Models)."
Write-Host "Done."
