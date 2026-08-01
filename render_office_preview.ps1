param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Output,
    [Parameter(Mandatory = $true)][ValidateSet('word', 'powerpoint')][string]$Kind
)

$ErrorActionPreference = 'Stop'
$application = $null
$document = $null

try {
    if ($Kind -eq 'word') {
        $application = New-Object -ComObject Word.Application
        $application.Visible = $false
        $document = $application.Documents.Open($Source, $false, $true)
        $document.ExportAsFixedFormat($Output, 17)
    }
    else {
        New-Item -ItemType Directory -Path $Output -Force | Out-Null
        $application = New-Object -ComObject PowerPoint.Application
        $document = $application.Presentations.Open($Source, $true, $false, $false)
        $number = 1
        foreach ($slide in $document.Slides) {
            $target = Join-Path $Output ("Slide{0}.PNG" -f $number)
            $slide.Export($target, "PNG", 1920, 1080)
            $number++
        }
    }
}
finally {
    if ($document) { $document.Close() }
    if ($application) { $application.Quit() }
}
