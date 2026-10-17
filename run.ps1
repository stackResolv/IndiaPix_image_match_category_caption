<#
.SYNOPSIS
    Windows PowerShell launcher for the photo matching + categorising pipeline.

.DESCRIPTION
    A PowerShell counterpart to run.bat. Both call the same run.py, so behaviour
    is identical; this one exists because PowerShell is the default terminal in
    VS Code on Windows and some people never open cmd.exe.

    What it fixes for PowerShell specifically:
      * Execution policy — run.bat is often blocked, this is a .ps1 you can run
        with a one-time -ExecutionPolicy Bypass (see below).
      * UTF-8 console — sets the code page and $OutputEncoding so non-ASCII
        filenames and captions do not turn into "?" or crash.
      * Console width — tells Python the real column count so the narration
        wraps correctly instead of running off the edge.

.EXAMPLE
    .\run.ps1
    # default shoot "_sort images"

.EXAMPLE
    .\run.ps1 input_folder
    # every shoot inside input_folder

.EXAMPLE
    .\run.ps1 input_folder --provider openai
    # use OpenAI for captions

.EXAMPLE
    .\run.ps1 "input_folder" "my_output"
    # batch + custom output base folder

.NOTES
    If PowerShell refuses to run this file with a script-execution message,
    that is the Execution Policy, not the tool. Either run it once with:
        powershell -ExecutionPolicy Bypass -File .\run.ps1 "input_folder"
    or unblock for the current session only:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

[CmdletBinding()]
param(
    # Everything after the known parameters is forwarded verbatim to run.py,
    # so new run.py flags never need a change here.
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $Args
)

$ErrorActionPreference = 'Stop'

# --- Force a UTF-8 console so non-ASCII paths/captions never crash ------------
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    # 65001 = UTF-8. Quiet: not every host lets you change this.
    $null = & chcp.com 65001 2>$null
}
catch {
    Write-Verbose "Could not switch the console to UTF-8; continuing anyway."
}
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'

# --- Find a Python interpreter (prefer the py launcher) ----------------------
$python = $null
$pythonArgs = @()

if (Get-Command py -ErrorAction SilentlyContinue) {
    $python = 'py'
    $pythonArgs = @('-3')
}
elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $python = 'python'
}
elseif (Get-Command python3 -ErrorAction SilentlyContinue) {
    $python = 'python3'
}

if (-not $python) {
    Write-Host ''
    Write-Host 'ERROR: Python was not found on PATH.' -ForegroundColor Red
    Write-Host 'Install it from https://www.python.org/downloads/ and tick'
    Write-Host '"Add Python to PATH", then reopen this window.'
    Write-Host ''
    exit 1
}

# --- Tell Python how wide this window really is ------------------------------
# run.py would otherwise guess 80 columns and wrap the narration badly.
try {
    $width = $Host.UI.RawUI.WindowSize.Width
    if ($width -ge 40) { $env:PIPELINE_WIDTH = "$width" }
}
catch {
    Write-Verbose "Could not read the window width; Python will detect it."
}

# --- Run the pipeline from the script's own folder --------------------------
Push-Location -LiteralPath $PSScriptRoot
try {
    & $python @pythonArgs 'run.py' @Args
    $code = $LASTEXITCODE
}
finally {
    Pop-Location
}

if ($code -ne 0) {
    Write-Host ''
    Write-Host 'ERROR: a step failed. See the messages above.' -ForegroundColor Red
    Write-Host ''
    exit $code
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
exit 0