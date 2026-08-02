param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Output,
    [Parameter(Mandatory = $true)][ValidateSet('word', 'powerpoint')][string]$Kind
)

$ErrorActionPreference = 'Stop'

# Office apps create work files under %TEMP%.  If TEMP/TMP are missing or point
# to an invalid path (e.g. a stale 8.3 short name), Word shows a modal dialog
# ("could not create the work file") and automation blocks forever.  Always
# normalise to the real per-user temp folder before starting Office.
$normalisedTemp = $null
foreach ($candidate in @($env:TEMP, $env:TMP, (Join-Path $env:LOCALAPPDATA 'Temp'), (Join-Path $env:USERPROFILE 'AppData\Local\Temp'))) {
    if ($candidate) {
        try {
            $resolved = [System.IO.Path]::GetFullPath([System.Environment]::ExpandEnvironmentVariables($candidate))
            if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
                New-Item -ItemType Directory -Path $resolved -Force | Out-Null
            }
            $probe = Join-Path $resolved ('fw-probe-' + [guid]::NewGuid().ToString('N') + '.tmp')
            Set-Content -LiteralPath $probe -Value 'x' -ErrorAction Stop
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            $normalisedTemp = $resolved
            break
        } catch {
            continue
        }
    }
}
if (-not $normalisedTemp) { throw 'No usable TEMP folder found for Office rendering.' }
$env:TEMP = $normalisedTemp
$env:TMP = $normalisedTemp

$application = $null
$document = $null
try {
    if ($Kind -eq 'word') {
        $application = New-Object -ComObject Word.Application
        $application.Visible = $false
        $application.DisplayAlerts = 0
        $document = $application.Documents.Open($Source, $false, $true)
        $document.ExportAsFixedFormat($Output, 17)
    }
    else {
        $application = New-Object -ComObject PowerPoint.Application
        $application.Visible = $false
        # Prefer a window-less open so no PowerPoint window flashes during a
        # background preview; fall back to WithWindow on hosts that refuse it.
        try {
            $document = $application.Presentations.Open($Source, $true, $false, $false)
        } catch {
            $document = $application.Presentations.Open($Source, $true, $false, $true)
        }
        # ExportAsFixedFormat is unreliable through PowerShell COM binding;
        # SaveAs with ppSaveAsPDF (32) produces the same native PDF.
        $document.SaveAs($Output, 32)
    }
    if (-not (Test-Path -LiteralPath $Output)) {
        throw "Export produced no output file: $Output"
    }
}
finally {
    if ($document) { try { $document.Close($false) } catch { try { $document.Close() } catch {} } }
    if ($application) { try { $application.Quit() } catch {} }
}