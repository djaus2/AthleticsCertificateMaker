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
h2 { font-size: 13pt; margin: 16pt 0 6pt; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #888; padding: 6px; text-align: left; }
th { background: #e8e8e8; }
</style>
'@

    try {
        $ResultGroups = [System.Collections.Generic.List[object]]::new()
        $CurrentGroup = [System.Collections.Generic.List[object]]::new()
        $CurrentHeading = $null

        foreach ($Result in $Data) {
            $Heading = [string]$Result.Heading
            if ($CurrentGroup.Count -gt 0 -and $Heading -cne $CurrentHeading) {
                $ResultGroups.Add([pscustomobject]@{
                    Heading = $CurrentHeading
                    Results = $CurrentGroup.ToArray()
                })
                $CurrentGroup = [System.Collections.Generic.List[object]]::new()
            }

            if ($CurrentGroup.Count -eq 0) {
                $CurrentHeading = $Heading
            }
            $CurrentGroup.Add($Result)
        }

        if ($CurrentGroup.Count -gt 0) {
            $ResultGroups.Add([pscustomobject]@{
                Heading = $CurrentHeading
                Results = $CurrentGroup.ToArray()
            })
        }

        $HtmlBody = [System.Collections.Generic.List[string]]::new()
        $HtmlBody.Add("<h1>Official Event Results</h1>")
        foreach ($Group in $ResultGroups) {
            if (-not [string]::IsNullOrWhiteSpace($Group.Heading)) {
                $SafeHeading = [System.Net.WebUtility]::HtmlEncode($Group.Heading)
                $HtmlBody.Add("<h2>$SafeHeading</h2>")
            }

            $TableHtml = $Group.Results |
                ConvertTo-Html -Fragment -Property No, Name, Distance, Time
            $HtmlBody.Add(($TableHtml -join [Environment]::NewLine))
        }

        $Html = ConvertTo-Html -Title "Official Event Results" -Head $Style -Body $HtmlBody.ToArray()
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