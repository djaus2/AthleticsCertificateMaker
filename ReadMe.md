# Athletics Certificate Generator

> Recently Aberfeldie Masters Athletics ran a 1 Hour Track Run. This project was created to generate certificates for participants and email them with the event results.
>
> Sample certificate:
>
> [004-Fred Nurk.png](https://github.com/djaus2/AthleticsCertificateMaker/blob/main/Output/PNG/004-Fred%20Nurk.png)

# Updates
- As a 1 hour event, it is assumed that all athletes complete the 1 hour time. Where a time has been added for a participant, `GenerateCertificates.ps1` does not generate a certificate and does not send anything to those participants but they are included in the results along with their time. The script `GenerateCertificatesTIME.ps1` generates an alternative certificate including their time for those athletes and sends it along with the results; this is only for those with a time.

This project generates:

- Participant certificate PNG files
- Event results PDF
- Outlook email drafts with appropriate attachments

The workflow is designed for the Aberfeldie Masters Athletics 1 Hour Track Run but can be adapted for similar events.

---

# Safe Approach (Recommended)

When testing changes or running the scripts for the first time, it is strongly recommended that emails are created as Outlook drafts rather than being sent automatically.

## Configure Draft-Only Behaviour

Comment out any:

```powershell
$Mail.Send()
```

and ensure any:

```powershell
$Mail.Save()
```

lines are enabled.

Example:

```powershell
#$Mail.Send()

$Mail.Save()
```

Result:

- Emails are created in the Outlook Drafts folder.
- No emails are sent automatically.
- Each draft can be reviewed individually before sending.
- Attachments can be verified before delivery.
- Subject lines and recipients can be checked manually.

This is the safest way to test any script modifications.

---

## Draft vs Send (Testing vs Live)

By default the scripts save generated messages to Outlook's Drafts folder for review.

### Drafts Only

```powershell
$Mail.Save()
```

Creates the email in:

```text
Outlook → Drafts
```

No email is sent.

### Live Sending

```powershell
$Mail.Send()
```

Immediately sends the email using the configured Outlook account.

### Sending Account

The sending account is specified near the top of each script:

```powershell
$SendUsingAccount = "account@location.com.au"
```

The value is resolved once when the script starts:

- If the account exists in the Outlook profile it is used for all emails.
- If it does not exist (for example the placeholder above), a warning is
  displayed and the default Outlook account is used instead.

### Startup Email-Mode Check

When each script starts it scans its own source to determine whether
`$Mail.Save()` and/or `$Mail.Send()` are enabled — commented-out lines are
ignored:

| State | Behaviour |
|-------|-----------|
| `Save` only | Runs silently (draft mode) |
| `Send` only | Warning popup: emails will actually be sent; must confirm to continue |
| Both enabled | Error popup and abort: only one may be enabled |
| Neither enabled | Warning popup that no emails will be produced |

---

# Overview

The process consists of:

1. Preparing the certificate background image.
2. Maintaining the participant spreadsheet.
3. Running the appropriate certificate generation script.
4. Reviewing generated certificates and email drafts.
5. Sending approved emails from Outlook.

Two certificate-generation workflows are available:

### GenerateCertificates.ps1

Creates standard distance certificates.

### GenerateCertificatesTIME.ps1

Creates certificates containing both distance and recorded time.

---

# Output Structure

```text
Output
├── PNG
├── PNG_TIME
└── Results
```

Where:

```text
PNG
```

contains standard certificates.

```text
PNG_TIME
```

contains certificates that include recorded times.

```text
Results
```

contains the generated PDF results file.

---

# Sample Output

## Standard Certificate

```text
Name

Fred Nurk

Distance

10145 metres
```

The Name and Distance labels are part of the template image.

## Time Certificate

```text
Name

Julie Smith

Distance

8730 metres

Time

45:56
```

On the time certificate all labels are drawn by the script: each label
sits above its value and the three groups are equally spaced.

---

# Creating the Certificate Background

The certificate background is based on a photograph of an athlete's singlet.

## Source Image

Use a high-resolution image.

The image should:

- Be centred on the page
- Occupy most of the certificate
- Be faded sufficiently so all text remains readable

## Recommended Settings

- White background
- Approximately 60% opacity
- Singlet centred
- No shadows

The Victorian Masters Athletics logo is positioned near the top of the page.

Two template images are included:

```text
CertificateTemplateorig.png
```

The original image with the *Name* and *Distance* labels baked in.
Used by `GenerateCertificates.ps1`.

```text
CertificateTemplate.png
```

The same image with those labels removed. Used by
`GenerateCertificatesTIME.ps1`, which draws all labels itself.

---

# Modifying Certificate Text

Certificate text is rendered dynamically by PowerShell using `DrawString()`.

## Standard Certificate

The *Name* and *Distance* labels are baked into `CertificateTemplateorig.png`.
The script draws the values centred beneath each label:

```text
Name

<Name>

Distance

<Distance> metres
```

## Time Certificate

`CertificateTemplate.png` contains no labels; the script draws all three
groups — label above value, equally spaced:

```text
Name

<Name>

Distance

<Distance> metres

Time

<Time>
```

---

# Adjusting Text Positions

Text locations are controlled using X/Y coordinates.

Example:

```powershell
$Graphics.DrawString(
    $Name,
    $NameFont,
    $Brush,
    $NameX,
    $NameY
)
```

Movement rules:

| Change | Action |
|----------|----------|
| Move down | Increase Y |
| Move up | Decrease Y |
| Move right | Increase X |
| Move left | Decrease X |

## Current Coordinate Values

### Standard Certificate

```powershell
$NameY = 1150
$DistanceY = 1380
```

### TIME Certificate

All strings are centred horizontally. Each label sits 80 px above its
value and the three groups are equally spaced:

```powershell
$NameLabelY     = 1030
$NameY          = 1110
$DistanceLabelY = 1190
$DistanceY      = 1270
$TimeLabelY     = 1350
$TimeY          = 1430
```

---

# Adjusting Fonts

Fonts are defined within the scripts.

Example:

```powershell
$NameFont = New-Object System.Drawing.Font(
    "Arial",
    20,
    [System.Drawing.FontStyle]::Bold
)
```

Common adjustments:

- Font family
- Size
- Bold
- Italic

---

# Spreadsheet Format

The spreadsheet must contain the following columns.

| Column | Required |
|----------|----------|
| No | Yes |
| Name | Yes |
| Email | Yes |
| Distance | Yes |
| Time | No |

Example:

| No | Name | Email | Distance | Time |
|----|------|--------|----------|------|
| 1 | Mary Swan | example@example.com | DNS | |
| 2 | Fred Nurk | example@example.com | 8730 | |
| 3 | John Smith | example@example.com | 12000 | 43:00 |

---

# Processing Rules

## GenerateCertificates.ps1

### DNS Entry

Example:

```text
Distance = DNS
```

Result:

- No certificate generated
- Results PDF attached
- Email draft created
- Participant remains in results PDF

### Standard Distance Result

Example:

```text
Distance = 8730
Time =
```

Result:

- Certificate generated
- Results PDF attached
- Certificate attached
- Email draft created

### Timed Result

Example:

```text
Distance = 12000
Time = 1:00:00
```

Result:

- No certificate generated
- No email generated
- Participant remains in results PDF

---

## GenerateCertificatesTIME.ps1

Processes only competitors with a value in the Time column.

### No Time Recorded

Example:

```text
Distance = 8730
Time =
```

Result:

- Skipped

### Time Recorded

Example:

```text
Distance = 12000
Time = 1:00:00
```

Result:

- Certificate generated
- Results PDF attached
- Certificate attached
- Email draft created

Certificates are written to:

```text
Output\PNG_TIME
```

---

# Running the Scripts

Open PowerShell and change to the repository folder.

## Standard Certificates

```powershell
cd C:\temp\AthleticsCertificateMaker

.\GenerateCertificates.ps1
```

## Time Certificates

```powershell
cd C:\temp\AthleticsCertificateMaker

.\GenerateCertificatesTIME.ps1
```

---

# Reviewing Output

### Standard Certificates

```text
Output\PNG
```

### Time Certificates

```text
Output\PNG_TIME
```

### Results PDF

```text
Results
```

Review all generated files and drafts before sending.

---

# Cleaning Generated Output

`clean.ps1` empties all generated content while leaving the folders in
place:

- `Output\*` — generated certificates and PDFs (keeps the tracked sample
  `004-Fred Nurk.png`)
- `Results\*` — generated results workbook and PDF
- `logs\*` — generated logs
- `~$*.xlsx` — Excel lock files
- `,\` — stray folder left by a buggy run (removed entirely)

Preview what would be deleted without removing anything:

```powershell
.\clean.ps1 -WhatIf
```

Then run for real:

```powershell
.\clean.ps1
```

---

# Outlook Requirements

## Outlook COM Automation

The scripts use:

```powershell
New-Object -ComObject Outlook.Application
```

The solution was tested using Outlook 2016.

A **classic (desktop) Outlook** profile is required — the "New Outlook"
app does not support COM automation and does not create a mail profile.

The profile must be set as the default so automation can open it without
prompting:

```text
Control Panel → Mail (Microsoft Outlook) → Show Profiles →
Always use this profile
```

Potential issues include:

- No mail profile configured ("no profile" COM error)
- Profile selection prompts
- Incorrect Outlook profile opening
- COM automation hanging

If issues occur:

```powershell
Get-Process Outlook -ErrorAction SilentlyContinue |
    Stop-Process -Force
```

Then restart Outlook and rerun the script.

---

## Sending Account

Configured near the top of each script:

```powershell
$SendUsingAccount = "account@location.com.au"
```

This value must match an account configured in Outlook. If it does not,
a warning is displayed and the default Outlook account is used.

---

# Excel / Printer Requirement

Although nothing is printed, Excel PDF generation may require a valid printer driver.

Recommended:

```text
Microsoft Print to PDF
```

or any installed printer.

Symptoms of a missing printer include:

- PDF generation failures
- Export errors
- Worksheet rendering issues

---

# Known Caveats


> Nb: It was found that the local printer needed to be turned on athough no actual printing was done.

> Depends upon classic desktop Outlook (2016) with a default mail profile configured.

---
---

# To Do

Parameterise these scripts so that this package can have simple reuse for slightly different contexts.
