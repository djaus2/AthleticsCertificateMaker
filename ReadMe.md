# Athletics Certificate Generator

> Recently Aberfeldie Masters Athletics ran a 1 Hour Track Run. This project was created to generate certificates for participants and email them with the event results.
>
> [Sample certiticate](https://github.com/djaus2/AthsCertificateMaker/blob/main/Output/PNG/004-Fred%20Nurk.png)

This project generates:

- Participant certificate PNG files
- Event results PDF
- Outlook email drafts with appropriate attachments

The workflow is designed for the Aberfeldie Masters Athletics 1 Hour Track Run but can be adapted for similar events.

---

## Draft vs Send (testing vs live)

By default the script saves generated messages to Outlook's Drafts folder for review. Key points:

- Mail.Save() saves the message to Drafts only (safe for testing).
- Mail.Send() sends the message immediately using the configured sending account; this is live and will deliver emails.
- The sending account ($SendUsingAccount) must match an account in Outlook.Session.Accounts or Outlook may fall back to the default account.
- Sending can trigger profile selection prompts or other Outlook UI depending on your configuration.
- If you wish to both keep a draft and send automatically you can call Save() then Send(), but verify that behaviour in your Outlook profile first.

To switch from draft-only to auto-send, uncomment the `# $Mail.Send()` lines in GenerateCertificates.ps1 (there are two locations where emails are created).


# Overview

The process consists of:

1. Preparing the certificate background image
2. Maintaining the participant spreadsheet
3. Running `GenerateCertificates.ps1`
4. Reviewing the generated files
5. Sending emails via Outlook

Output files are created in:

```text
Output\
├── PNG\
├── PDF\
└── Results\
```

---

## Sample Output

```text
PS C:\Certificates> .\GenerateCertificates.ps1
Loading participants...
Participants found: 3
Creating PDF results file...
Results PDF created.
Outlook COM created successfully

Row:
  Name     = 'Participant One'
  Email    = 'account@location.com.au'
  Distance = 'DNS'
  Time     = ''
Processing DNS for Participant One
DNS: Participant One - results only

=== RESULTS EMAIL ===
Name: Participant One
Email: account@location.com.au
Results PDF: C:\Certificates\Results\Aberfeldie-1-Hour-Track-Run-Results-2026.pdf
Results exists: True
Not adding certificate attachment...
Adding results attachment...
Number of attachments: 1
Saving results email draft...
Results email draft saved.
Results email sent.
Results email processing complete.

Row:
  Name     = 'Participant Two'
  Email    = 'account@location.com.au'
  Distance = '8730'
  Time     = '45:56'
Skipping timed result: Participant Two (45:56)

Row:
  Name     = 'Fred Nurk'
  Email    = 'account@location.com.au'
  Distance = '1145'
  Time     = ''
Processing DNS for Fred Nurk
Generating certificate for Fred Nurk
Created C:\Certificates\Output\PNG\004-Fred Nurk.png
Number of attachments: 2
Saving certificate email draft with cert and pdf...
Results email sent with certs and pdfs.

Completed.
Certificates: C:\Certificates\Output\PNG
Results PDF:  C:\Certificates\Results\Aberfeldie-1-Hour-Track-Run-Results-2026.pdf
Email drafts created in Outlook.
```

# Creating the Certificate Background

The certificate background is based on a photograph of an athlete's singlet.

## Source Image

Use a high-resolution photo of the singlet.

The image should:

- Be centred on the certificate
- Occupy most of the page
- Be faded so certificate text remains readable

## Background Settings

Recommended settings:

- Opacity: approximately 60%
- White background
- Singlet centred on page
- No shadows or colour shifts

The Victorian Masters Athletics logo is positioned near the top of the certificate and scaled to approximately 200% of its original size.

The background image is stored as:

```text
CertificateTemplate.png
```

---

# Modifying Certificate Text

All certificate text is added by the PowerShell script.

Common fields include:

```text
Presented to
<Name>

for participating in the

Aberfeldie Masters Athletics
1 Hour Track Run

Distance Achieved

<Distance> metres
```

## Adjusting Text Position

Text placement is controlled by X/Y coordinates in the script.

Example:

```powershell
$Graphics.DrawString(
    $Participant.Name,
    $NameFont,
    $Brush,
    500,
    750,
    $CenterFormat
)
```

Where:

- First coordinate = X position
- Second coordinate = Y position

Moving text:

| Change | Action |
|---------|---------|
| Move down | Increase Y |
| Move up | Decrease Y |
| Move right | Increase X |
| Move left | Decrease X |

---

# Adjusting Fonts

Fonts are defined within the script.

Example:

```powershell
$NameFont = New-Object System.Drawing.Font(
    "Arial",
    42,
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

The spreadsheet must contain the following columns:

| Column | Required |
|----------|----------|
| No | Yes |
| Name | Yes |
| Email | Yes |
| Distance | Yes |
| Time | Yes |

Example:

| No | Name | Email | Distance | Time |
|----|------|--------|----------|------|
| 1 | Maria Abfalter | example@example.com | DNS | |
| 2 | Fred Nurk | example@example.com | 8730 | 45:56 |
| 3 | John Smith | example@example.com | 12000 | 1:00:00 |

---

# Processing Rules

## DNS Entries

If:

```text
Distance = DNS
```

Then:

- No certificate generated
- Results PDF emailed
- Participant remains in results PDF

---

## Completed Time

Example:

```text
Distance = 8730
Time = <is blank>
```

Then:

- Certificate generated
- Results PDF attached
- Certificate attached
- Email created

---

## Did not complete 1 hr

Example:

```text
Distance = 12000
Time = 40:35
```

Then:

- No certificate generated
- No email generated
- Participant remains in results PDF

This behaviour was implemented to allow later addition of time results for participants who did not complete the 1 hour run.

---

# Running the Script

Open PowerShell:

```powershell
cd C:\Certificates
.\GenerateCertificates.ps1
```

Expected output includes:

```text
Loading participants...
Participants found: xx

Creating PDF results file...
Results PDF created.

Creating certificates...
Creating Outlook drafts...
```

---

# Reviewing Output

Certificate PNG files:

```text
Output\PNG
```

Results PDF:

```text
Output\PDF
```

Review all generated files before sending emails.

---

# Outlook Requirements

## Outlook 2016

The email generation component currently relies on Outlook COM automation.

The script was tested using:

```text
Microsoft Outlook 2016
```

Other Outlook installations may behave differently.

Possible issues include:

- Profile selection prompts
- Outlook opening the wrong profile
- COM automation hanging when Outlook is already running

If issues occur:

1. Close Outlook.
2. End all Outlook processes.
3. Run the script again.

Example:

```powershell
Get-Process Outlook -ErrorAction SilentlyContinue |
    Stop-Process -Force

## Sending account

The script now exposes the sending account as a top-level variable in GenerateCertificates.ps1:

```powershell
$SendUsingAccount = "account@location.com.au"
```

Edit that value to match the email account you want Outlook to send from (for example `"onehour@sportronics.com.au"`). The value must match an account present in your Outlook profile (Outlook.Session.Accounts).

```text


---

# Excel / Printer Requirement

Although the script does not print anything, Excel may require a valid printer driver to be available.

Symptoms include:

- Excel opening but failing during export
- PDF generation errors
- Worksheet rendering problems

Recommended:

- Ensure at least one printer is installed.
- A PDF printer is sufficient.
- Microsoft Print to PDF is recommended.

---

# Known Caveats

## Outlook COM Automation

The script uses:

```powershell
New-Object -ComObject Outlook.Application
```

If Outlook profiles become corrupted or multiple Outlook versions are installed, email creation may fail.

---

## Spreadsheet File Locks

Close the spreadsheet before running the script.

Excel locking the workbook can prevent:

- Reading participant data
- Generating output files

---

## Email Review

It is strongly recommended to:

- Generate drafts first
- Review recipients
- Verify certificate attachments
- Verify PDF attachment

before sending.

---

# Recommended Folder Structure

```text
C:\Certificates
│
├── GenerateCertificates.ps1
├── SendCertificates.ps1
├── Participants.xlsx
├── CertificateTemplate.png
├── VMA_Logo.png
│
└── Output
    ├── PNG
    ├── PDF
    └── Results
```

---

Note: GenerateCertificates.ps1 creates and (optionally) sends emails as part of its workflow. Because of this, SendCertificates.ps1 is not normally required and has not been tested as part of this repository's primary flow. Use SendCertificates.ps1 only if you have a specific separate sending workflow and verify its behaviour before relying on it.


# Cleaning generated files

A helper script, `clean.ps1`, is included to remove files that are generated by the scripts (for example files under Output\ and other script-created or editor temporary files that are listed in .gitignore).

Important points:

- The script deletes only files that Git reports as "ignored" (via `git ls-files --others --ignored --exclude-standard`). It will not remove tracked repository source files.
- The script runs from the repository root and only deletes paths under the repository root.
- The script explicitly skips anything inside the `.vs` folder.

Usage:

```powershell
.\clean.ps1 -WhatIf   # preview files that would be removed
.\clean.ps1           # actually remove ignored/generated files
```

Run the preview first to confirm what will be deleted.

# Revision History

## Version 1.0

Features:

- PNG certificate generation
- Results PDF generation
- Outlook email draft generation
- DNS results-only handling
- Timed-result exclusion handling
- Automated attachment processing