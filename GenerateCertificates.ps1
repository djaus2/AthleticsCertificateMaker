$CertificateTextDefinitionFile = "CertificateText.ps1"

Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

#
# Check which email actions are enabled in this script:
#   $Mail.Save()  - save to Drafts only
#   $Mail.Send()  - actually send
# Both enabled    -> abort: only one can be enabled
# Send enabled    -> warn and confirm before continuing
# Neither enabled -> warn that no certificate emails will be produced
#
Add-Type -AssemblyName PresentationFramework

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $PSCommandPath,
    [ref]$tokens,
    [ref]$parseErrors
)

$commentExtents = @(
    $tokens |
        Where-Object { $_.Kind -eq 'Comment' } |
        ForEach-Object { $_.Extent }
)

function Test-MailCallEnabled {
    param([string]$MethodName)

    $calls = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Member.Value -eq $MethodName -and
            $node.Expression.Extent.Text -eq '$Mail'
    }, $true)

    foreach ($call in $calls) {
        $insideComment = $false
        foreach ($extent in $commentExtents) {
            if ($call.Extent.StartOffset -ge $extent.StartOffset -and
                $call.Extent.EndOffset -le $extent.EndOffset) {
                $insideComment = $true
                break
            }
        }
        if (-not $insideComment) {
            return $true
        }
    }
    return $false
}

$SendEnabled = Test-MailCallEnabled 'Send'
$SaveEnabled = Test-MailCallEnabled 'Save'

if ($SendEnabled -and $SaveEnabled) {
    [System.Windows.MessageBox]::Show(
        "Both `$Mail.Save() and `$Mail.Send() are enabled.`nOnly one can be enabled.`n`nThe script will now abort.",
        "Email configuration error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
    exit 1
}

if ($SendEnabled) {
    $result = [System.Windows.MessageBox]::Show(
        "`$Mail.Send() is enabled.`n`nCertificate emails will actually be SENT.`n`nContinue?",
        "Send confirmation",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 1
    }
}

if (-not $SendEnabled -and -not $SaveEnabled) {
    $result = [System.Windows.MessageBox]::Show(
        "Neither `$Mail.Save() nor `$Mail.Send() is enabled.`n`nNo certificate emails will be created.`n`nContinue?",
        "Email disabled",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 1
    }
}

$RootFolder = (Get-Location).Path

# Header text definitions shared by both scripts (title, subtitle, host, date)
. (Join-Path $PSScriptRoot $CertificateTextDefinitionFile)
. (Join-Path $PSScriptRoot "Export-ResultsPdf.ps1")
$CsvFile = Join-Path $RootFolder $ParticipantCsvFile
$TemplateFile = Join-Path $RootFolder $CertificateTemplateFile

$PngFolder = Join-Path $RootFolder "Output\PNG"
$ResultsFolder = Join-Path $RootFolder "Results"

$ResultsCsv = Join-Path $ResultsFolder "OfficialResults.csv"
$ResultsPdf = Join-Path $ResultsFolder "Aberfeldie-1-Hour-Track-Run-Results-2026.pdf"

New-Item -ItemType Directory -Force -Path $PngFolder | Out-Null
New-Item -ItemType Directory -Force -Path $ResultsFolder | Out-Null

Write-Host "Loading participants..." -ForegroundColor Cyan

$Participants = Import-Csv $CsvFile

$Participants = $Participants | Where-Object {
    -not [string]::IsNullOrWhiteSpace($_.Name)
}

Write-Host "Participants found: $($Participants.Count)"

#
# Create public results files WITHOUT email addresses
#
$OfficialResults = @($Participants | Select-Object No, Name, Distance, Time)
$OfficialResults |
    Export-Csv $ResultsCsv -NoTypeInformation -Encoding UTF8

Write-Host "Creating PDF results file..." -ForegroundColor Cyan
Export-ResultsPdf -Data $OfficialResults -Path $ResultsPdf

Write-Host "Results PDF created." -ForegroundColor Green

#
# Outlook
#
$Outlook = New-Object -ComObject Outlook.Application

if ($null -eq $Outlook)
{
    throw "Failed to create Outlook COM object"
}

Write-Host "Outlook COM created successfully" -ForegroundColor Green

#
# Resolve send-as account once; fall back to default if not configured
#
$SendAccount = $Outlook.Session.Accounts.Item($SendUsingAccount)
if ($null -eq $SendAccount) {
    Write-Host "Account '$SendUsingAccount' not found - using default account" `
        -ForegroundColor Yellow
}

foreach ($Participant in $Participants)
{
    if ([string]::IsNullOrWhiteSpace($Participant.Name))
    {
        continue
    }

    $Number   = $Participant.No
    $Name     = [string]$Participant.Name
    $Email    = [string]$Participant.Email
    $Distance = [string]$Participant.Distance
    $Time     = [string]$Participant.Time


    #
    # Timed results
    # No certificate / no email
    #
    if (-not [string]::IsNullOrWhiteSpace($Time))
    {
        Write-Host "Skipping timed result: $Name ($Time)" `
            -ForegroundColor Yellow
        continue
    }

    Write-Host "Processing DNS for $Name" -ForegroundColor Cyan
    #
    # DNS
    # Results PDF only
    #
    if ($Distance.Trim().ToUpper() -eq "DNS")
    {
        Write-Host "DNS: $Name - results only" -ForegroundColor Cyan

        if (-not [string]::IsNullOrWhiteSpace($Email))
        {
            Write-Host ""
            Write-Host "=== RESULTS EMAIL ===" -ForegroundColor Cyan
            Write-Host "Name: $Name"
            Write-Host "Email: $Email"
            Write-Host "Results PDF: $ResultsPdf"
            Write-Host "Results exists: $(Test-Path $ResultsPdf)"

            $Mail = $Outlook.CreateItem(0)

            $Mail.SendUsingAccount = $SendAccount

            $Mail.To = $Email

            $Mail.Subject =
                "Aberfeldie Masters Athletics - 1 Hour Track Run Certificate"

            $Mail.HTMLBody = @"
    <p>Dear $Name,</p>

    <p>Please find attached:</p>

    <ul>
    <li>Your participation certificate</li>
    <li>The official event results</li>
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

            Write-Host "Not adding certificate attachment..."

            Write-Host "Adding results attachment..."
            $Mail.Attachments.Add($ResultsPdf) | Out-Null

            Write-Host "Number of attachments: $($Mail.Attachments.Count)" `
                -ForegroundColor DarkGreen  -BackgroundColor Yellow

            Write-Host "Saving results email draft..." `
                -ForegroundColor DarkGreen  -BackgroundColor Yellow

            # For testing: make changes to GenerateCertificates.ps1 then run.
            # Emails will be saved to the Drafts folder only.
            # To actually send automatically on generation uncomment the following line:
            #$Mail.Send()

            # Save the email to the draft folder without sending it
            $Mail.Save()

            Write-Host "Results email draft saved." `
                -ForegroundColor DarkGreen  -BackgroundColor Yellow

            Write-Host "Results email processing complete." `
                -ForegroundColor DarkGreen  -BackgroundColor Yellow
        }

        continue
    }

    Write-Host "Generating certificate for $Name"

    $Bitmap = New-Object System.Drawing.Bitmap $TemplateFile

    $Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)

    $Graphics.SmoothingMode =
        [System.Drawing.Drawing2D.SmoothingMode]$CertificateSmoothingMode

    $Graphics.TextRenderingHint =
        [System.Drawing.Text.TextRenderingHint]$CertificateTextRenderingHint

    $DefaultBrush = New-Object System.Drawing.SolidBrush(
        ([System.Drawing.Color]::FromName($DefaultTextColor))
    )

    #
    # Draw header text (defined in CertificateText.ps1)
    #
    foreach ($item in $CertificateText) {
        $ItemBrush = New-Object System.Drawing.SolidBrush(
            ([System.Drawing.Color]::FromName([string]$item.Color))
        )

        $font = New-Object System.Drawing.Font(
            $item.Font,
            $item.Size,
            [System.Drawing.FontStyle]($item.Style),
            [System.Drawing.GraphicsUnit]::Pixel
        )

        $textSize = $Graphics.MeasureString($item.Text, $font)
        $textX = ($Bitmap.Width - $textSize.Width) / 2

        $Graphics.DrawString(
            $item.Text,
            $font,
            $ItemBrush,
            $textX,
            $item.Y
        )

        $font.Dispose()
    $ItemBrush.Dispose()
    }

    #
    # Draw fields (defined in CertificateText.ps1)
    #
    foreach ($field in $CertificateFields.Standard) {
        $labelColorName = if ($field.LabelColor) { [string]$field.LabelColor } elseif ($DefaultLabelColor) { [string]$DefaultLabelColor } else { $DefaultTextColor }
        $valueColorName = if ($field.ValueColor) { [string]$field.ValueColor } elseif ($DefaultValueColor) { [string]$DefaultValueColor } else { $DefaultTextColor }

        $labelBrush = New-Object System.Drawing.SolidBrush(
            ([System.Drawing.Color]::FromName($labelColorName))
        )
        $valueBrush = New-Object System.Drawing.SolidBrush(
            ([System.Drawing.Color]::FromName($valueColorName))
        )

        $labelFont = New-Object System.Drawing.Font(
            $field.LabelFont,
            $field.LabelSize,
            [System.Drawing.FontStyle]($field.LabelStyle),
            [System.Drawing.GraphicsUnit]::Pixel
        )

        $valueFont = New-Object System.Drawing.Font(
            $field.ValueFont,
            $field.ValueSize,
            [System.Drawing.FontStyle]($field.ValueStyle),
            [System.Drawing.GraphicsUnit]::Pixel
        )

        $labelX = ($Bitmap.Width -
            $Graphics.MeasureString($field.Label, $labelFont).Width) / 2

        $value = "$($Participant.$($field.Name))$($field.Suffix)"
        $valueX = ($Bitmap.Width -
            $Graphics.MeasureString($value, $valueFont).Width) / 2

        $Graphics.DrawString(
            $field.Label,
            $labelFont,
            $labelBrush,
            $labelX,
            $field.LabelY
        )

        $Graphics.DrawString(
            $value,
            $valueFont,
            $valueBrush,
            $valueX,
            $field.ValueY
        )

        $labelFont.Dispose()
        $valueFont.Dispose()
        $labelBrush.Dispose()
        $valueBrush.Dispose()
    }

    $DefaultBrush.Dispose()

    $SafeName = $Name -replace '[\\/:*?"<>|]', ''

    $CertificateFile = Join-Path $PngFolder (
        "{0:D3}-{1}.png" -f [int]$Number, $SafeName
    )

    $Bitmap.Save(
        $CertificateFile,
        [System.Drawing.Imaging.ImageFormat]::Png
    )

    $Graphics.Dispose()
    $Bitmap.Dispose()

    Write-Host "Created $CertificateFile" `
        -ForegroundColor Green

    if (-not [string]::IsNullOrWhiteSpace($Email))
    {
        $Mail = $Outlook.CreateItem(0)

        $Mail.SendUsingAccount = $SendAccount

        $Mail.To = $Email

        $Mail.Subject =
            "Aberfeldie Masters Athletics - 1 Hour Track Run Certificate"

        $Mail.HTMLBody = @"
<p>Dear $Name,</p>

<p>Please find attached:</p>

<ul>
<li>Your participation certificate</li>
<li>The official event results</li>
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

        Write-Host "Number of attachments: $($Mail.Attachments.Count)" -ForegroundColor Yellow -BackgroundColor DarkGreen

        #
        # Draft for review
        Write-Host "Saving certificate email draft with cert and pdf..." -ForegroundColor Yellow -BackgroundColor DarkGreen
        # For testing: make changes to GenerateCertificates.ps1 then run.
        # Emails will be saved to the Drafts folder only.
        # To actually send automatically on generation uncomment the following line:
        #$Mail.Send()

        # Save the email to the draft folder without sending it
        $Mail.Save()

        Write-Host "Results email saved as draft with certs and pdfs." `
            -ForegroundColor Yellow -BackgroundColor DarkGreen
    }
}

Write-Host ""
Write-Host "Completed." -ForegroundColor Green
Write-Host "Certificates: $PngFolder"
Write-Host "Results PDF:  $ResultsPdf"
Write-Host "Email drafts created in Outlook."