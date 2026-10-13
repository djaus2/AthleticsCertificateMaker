function Export-ResultsPdf {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Data,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $EdgeCandidates = @(
        (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
    )
    $EdgePath = $EdgeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $EdgePath) {
        throw "Microsoft Edge is required to create the results PDF."
    }

    $TempFolder = Join-Path $env:TEMP ([guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
    $HtmlPath = Join-Path $TempFolder "results.html"
    $ProfilePath = Join-Path $TempFolder "edge-profile"

    $Style = @'
<style>
@page { size: A4; margin: 16mm; }
body { font-family: Arial, sans-serif; font-size: 10pt; color: #111; }
h1 { font-size: 18pt; margin: 0 0 12pt; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #888; padding: 6px; text-align: left; }
th { background: #e8e8e8; }
</style>
'@

    try {
        $Html = $Data |
            ConvertTo-Html -Property No, Name, Distance, Time `
                -Title "Official Event Results" `
                -Head $Style `
                -PreContent "<h1>Official Event Results</h1>"
        $Html | Set-Content -LiteralPath $HtmlPath -Encoding UTF8

        $FullPath = [System.IO.Path]::GetFullPath($Path)
        $HtmlUri = ([System.Uri]$HtmlPath).AbsoluteUri
        $Arguments = @(
            "--headless"
            "--disable-gpu"
            "--no-first-run"
            "--no-pdf-header-footer"
            "--user-data-dir=`"$ProfilePath`""
            "--print-to-pdf=`"$FullPath`""
            "`"$HtmlUri`""
        )

        $Process = Start-Process -FilePath $EdgePath -ArgumentList $Arguments -Wait -PassThru
        if ($Process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $FullPath)) {
            throw "Microsoft Edge failed to create the results PDF."
        }
    }
    finally {
        Remove-Item -LiteralPath $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}