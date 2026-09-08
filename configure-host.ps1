param(
    [Parameter(Mandatory = $true)]
    [string]$BaseUrl,

    [string]$Output = "manifest.custom.xml"
)

$ErrorActionPreference = 'Stop'

$base = $BaseUrl.TrimEnd('/')
if ($base -notmatch '^https://') {
    throw 'BaseUrl must start with https://'
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$template = Join-Path $root 'manifest.template.xml'
$outputPath = Join-Path $root $Output

$content = Get-Content -Raw -Path $template
$content = $content.Replace('__BASE_URL__', $base)
Set-Content -Path $outputPath -Value $content -Encoding UTF8

Write-Host "Created: $outputPath"
Write-Host "Task pane URL: $base/taskpane.html"
