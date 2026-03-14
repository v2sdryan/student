$ErrorActionPreference = "Stop"

$edgePathCandidates = @(
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)

$edgePath = $edgePathCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edgePath) {
    throw "Microsoft Edge was not found."
}

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$buildDir = Join-Path $baseDir ".pdf-build"

if (-not (Test-Path $buildDir)) {
    New-Item -ItemType Directory -Path $buildDir | Out-Null
}

Add-Type -AssemblyName System.Web

Get-ChildItem -Path $baseDir -Filter "*.txt" | Sort-Object Name | ForEach-Object {
    $textPath = $_.FullName
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    $htmlPath = Join-Path $buildDir ($stem + ".html")
    $pdfPath = Join-Path $baseDir ($stem + ".pdf")

    $rawText = Get-Content -Path $textPath -Raw -Encoding UTF8
    $safeText = [System.Web.HttpUtility]::HtmlEncode($rawText)
    $title = [System.Web.HttpUtility]::HtmlEncode($stem)

    $html = @"
<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <title>$title</title>
    <style>
        @page {
            size: A4;
            margin: 16mm 14mm 18mm;
        }

        body {
            margin: 0;
            font-family: "Microsoft JhengHei", "Noto Sans TC", sans-serif;
            color: #153147;
        }

        h1 {
            margin: 0 0 14px;
            font-size: 22px;
            color: #2c7faa;
        }

        pre {
            margin: 0;
            white-space: pre-wrap;
            word-break: break-word;
            line-height: 1.8;
            font-size: 13.5px;
            font-family: inherit;
        }
    </style>
</head>
<body>
    <h1>$title</h1>
    <pre>$safeText</pre>
</body>
</html>
"@

    Set-Content -Path $htmlPath -Value $html -Encoding UTF8

    $fileUri = "file:///" + ($htmlPath -replace "\\", "/")
    & $edgePath --headless --disable-gpu --print-to-pdf="$pdfPath" --print-to-pdf-no-header $fileUri | Out-Null

    if (-not (Test-Path $pdfPath)) {
        throw "Failed to generate PDF for $($_.Name)"
    }
}
