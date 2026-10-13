$BaseFolder = "C:\Certificates"
. (Join-Path $PSScriptRoot "CertificateText.ps1")

Write-Host ""
Write-Host "=== Certificate System Validation ===" -ForegroundColor Cyan
Write-Host ""

$Errors = 0

function Test-ItemExists {
    param(
        [string]$Path,
        [string]$Description
    )

    if (Test-Path $Path) {
        Write-Host "[OK] $Description" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Missing $Description" -ForegroundColor Red
        $script:Errors++
    }
}

# Required files
$CsvFile = Join-Path $BaseFolder $ParticipantCsvFile
$TemplateFile = Join-Path $BaseFolder $CertificateTemplateFile
$ScriptFile = Join-Path $BaseFolder "Send-Certificates.ps1"

Test-ItemExists $CsvFile "Participant CSV"
Test-ItemExists $TemplateFile "Certificate template PNG"
Test-ItemExists $ScriptFile "Certificate mail script"

$EdgeCandidates = @(
    (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
    (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe")
)
if ($EdgeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1)
{
    Write-Host "[OK] Microsoft Edge available for PDF generation" -ForegroundColor Green
}
else
{
    Write-Host "[ERROR] Microsoft Edge is required for PDF generation" -ForegroundColor Red
    $Errors++
}

# Required folders
$OutputFolder = Join-Path $BaseFolder "Output"
$PngFolder = Join-Path $OutputFolder "PNG"
$PdfFolder = Join-Path $OutputFolder "PDF"
$LogsFolder = Join-Path $BaseFolder "Logs"

foreach ($Folder in @($OutputFolder, $PngFolder, $PdfFolder, $LogsFolder))
{
    if (Test-Path $Folder)
    {
        Write-Host "[OK] Folder exists: $Folder" -ForegroundColor Green
    }
    else
    {
        Write-Host "[WARN] Creating folder: $Folder" -ForegroundColor Yellow
        New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    }
}

Write-Host ""

# Validate participant CSV
if (Test-Path $CsvFile)
{
    try
    {
        $Data = @(Import-Csv $CsvFile -ErrorAction Stop)

        if ($Data.Count -eq 0)
        {
            Write-Host "[ERROR] CSV contains no rows" -ForegroundColor Red
            $Errors++
        }
        else
        {
            Write-Host "[OK] CSV contains $($Data.Count) entrants" -ForegroundColor Green
        }

        $FirstRow = $Data | Select-Object -First 1

        $Columns = $FirstRow.PSObject.Properties.Name

        $RequiredColumns = @(
            "No",
            "Name",
            "Email",
            "Distance"
        )

        foreach ($Column in $RequiredColumns)
        {
            if ($Columns -contains $Column)
            {
                Write-Host "[OK] Column found: $Column" -ForegroundColor Green
            }
            else
            {
                Write-Host "[ERROR] Missing column: $Column" -ForegroundColor Red
                $Errors++
            }
        }
    }
    catch
    {
        Write-Host "[ERROR] Unable to read participant CSV"
        Write-Host $_.Exception.Message
        $Errors++
    }
}

Write-Host ""

# Validate PNG
if (Test-Path $TemplateFile)
{
    try
    {
        Add-Type -AssemblyName System.Drawing

        $Img = [System.Drawing.Image]::FromFile($TemplateFile)

        Write-Host "[OK] PNG readable" -ForegroundColor Green
        Write-Host "     Size: $($Img.Width) x $($Img.Height)"

        if ($Img.Width -lt 2000)
        {
            Write-Host "[WARN] PNG may be low resolution for printing" -ForegroundColor Yellow
        }

        $Img.Dispose()
    }
    catch
    {
        Write-Host "[ERROR] Cannot open PNG file" -ForegroundColor Red
        $Errors++
    }
}

Write-Host ""
Write-Host "===================================="

if ($Errors -eq 0)
{
    Write-Host "VALIDATION PASSED" -ForegroundColor Green
}
else
{
    Write-Host "VALIDATION FAILED ($Errors errors)" -ForegroundColor Red
}

Write-Host "===================================="