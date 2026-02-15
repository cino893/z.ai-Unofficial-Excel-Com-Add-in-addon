# build.ps1 - PowerShell 7 wrapper to build Z.AI Excel Add-in
# Usage: pwsh -File build.ps1

Param()

# Ensure running under PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "PowerShell 7+ is recommended. Current version: $($PSVersionTable.PSVersion)"
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Using script directory: $scriptDir"

$buildVbs = Join-Path $scriptDir 'build.vbs'
if (-not (Test-Path $buildVbs)) {
    Write-Error "build.vbs not found in $scriptDir"
    exit 1
}

Write-Host "Launching VBScript builder (cscript)..."
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = 'cscript.exe'
$psi.Arguments = "//nologo `"$buildVbs`""
$psi.WorkingDirectory = $scriptDir
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false

$proc = [System.Diagnostics.Process]::Start($psi)
$stdout = $proc.StandardOutput.ReadToEnd()
$stderr = $proc.StandardError.ReadToEnd()
$proc.WaitForExit()

if ($stdout) { Write-Host $stdout }
if ($stderr) { Write-Host $stderr -ForegroundColor Red }
exit $proc.ExitCode
