Import-Module ImportExcel

$RootFolder = "C:\Certificates"

$ExcelFile = Join-Path $RootFolder "Resultsentrants-aberfeldie-one-hour-track-challenge.xlsx"

$PngFolder = Join-Path $RootFolder "Output\PNG"

$ResultsPdf = Join-Path `
    $RootFolder `
    "Results\Aberfeldie-1-Hour-Track-Run-Results-2026.pdf"

if (-not (Test-Path $ResultsPdf))
{
    Write-Error "Results PDF not found: $ResultsPdf"
    exit
}

$Participants = Import-Excel $ExcelFile

$Outlook = New-Object -ComObject Outlook.Application

foreach ($Participant in $Participants)
{
    $Number   = $Participant.No
    $Name     = [string]$Participant.Name
    $Email    = [string]$Participant.'Email Address'
    $Distance = [string]$Participant.Distance
    $Time     = [string]$Participant.Time

    if (:IsNullOrWhiteSpace($Name))
    {
        continue
    }

    if (:IsNullOrWhiteSpace($Email))
    {
        Write-Host "Skipping $Name - no email address" `
            -ForegroundColor Yellow
        continue
    }

    #
    # Timed result
    # Send nothing
    #
    if (-not :IsNullOrWhiteSpace($Time))
    {
        Write-Host "Skipping $Name - timed result ($Time)" `
            -ForegroundColor Yellow

        continue
    }

    #
    # DNS
    # Results only
    #
    if ($Distance.Trim().ToUpper() -eq "DNS")
    {
        Write-Host "Sending results only to $Name" `
            -ForegroundColor Cyan

        $Mail = $Outlook.CreateItem(0)

        $Mail.To = $Email

        $Mail.Subject =
            "Aberfeldie Masters Athletics - 1 Hour Track Run Results"

        $Mail.HTMLBody = @"
<p>Dear $Name,</p>

<p>Please find attached the official event results.</p>

<p>Regards<br/>
Aberfeldie Masters Athletics</p>
"@

        $Mail.Attachments.Add($ResultsPdf) | Out-Null

        $Mail.Send()

        continue
    }

    #
    # Certificate recipient
    #
    $SafeName = $Name -replace '[\\/:*?"<>|]', ''

    $CertificateFile = Join-Path $PngFolder (
        "{0:D3}-{1}.png" -f [int]$Number, $SafeName
    )

    if (-not (Test-Path $CertificateFile))
    {
        Write-Host "Certificate missing for $Name" `
            -ForegroundColor Red
        continue
    }

    Write-Host "Sending certificate to $Name" `
        -ForegroundColor Green

    $Mail = $Outlook.CreateItem(0)

    $Mail.To = $Email

    $Mail.Subject =
        "Aberfeldie Masters Athletics - 1 Hour Track Run Certificate"

    $Mail.HTMLBody = @"
<p>Dear $Name,</p>

<p>Please find attached:</p>

<ul>
<li>Your participation certificate</li>
<li>The official results</li>
</ul>

<p>
Recorded distance:
<strong>$Distance metres</strong>
</p>

<p>Thank you for participating.</p>

<p>
Regards<br/>
Aberfeldie Masters Athletics
</p>
"@

    $Mail.Attachments.Add($CertificateFile) | Out-Null
    $Mail.Attachments.Add($ResultsPdf) | Out-Null

    $Mail.Send()
}

Write-Host ""
Write-Host "Email processing complete." `
    -ForegroundColor Green