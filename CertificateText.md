# CertificateText.ps1 — Certificate Text Configuration

`CertificateText.ps1` is a shared configuration file used by the certificate
generation and mail scripts. Certificate wording, presentation, input
spreadsheet filename, and sending account are defined here.

To change wording, fonts or positions, edit this file once and both
scripts pick it up on the next run.

## Certificate template image

Set `$CertificateTemplateFile` to the PNG filename in the certificate
folder. Both generation scripts and `Check.ps1` use this setting. The
participant spreadsheet filename is shared by the generation scripts,
`Check.ps1`, and `Send-Certificates.ps1`. `$SendUsingAccount` configures the
sender account for both generation scripts; `Send-Certificates.ps1` continues
to use Outlook's default account.

```powershell
$CertificateTemplateFile = "CertificateTemplate.png"
$ParticipantSpreadsheetFile = "results.xlsx"
$SendUsingAccount = "account@location.com.au"
$CertificateTextColor = @(0, 20, 90) # RGB
$CertificateSmoothingMode = "HighQuality"
$CertificateTextRenderingHint = "AntiAliasGridFit"
```

The text color and rendering settings are shared by both generation scripts.

---

## Header text — `$CertificateText`

The four lines at the top of the certificate. Each is drawn centred
horizontally at `Y`, at its natural width (no scaling).

| Name | Text | Font | Size | Style | Y |
|------|------|------|------|-------|---|
| Title | 1 HOUR TRACK RUN | Times New Roman | 77 px | Bold | 70 |
| SubTitle | Moonee Valley Athletics Centre | Arial Narrow | 51 px | Regular | 180 |
| Host | Aberfeldie Masters Athletics | Times New Roman | 38 px | Italic | 247 |
| Date | 16 September 2026 | Arial Narrow | 34 px | Regular | 312 |

Each entry is a hashtable:

```powershell
@{
    Name  = "Title"                     # identifier only
    Text  = "1 HOUR TRACK RUN"          # what gets drawn
    Font  = "Times New Roman"           # font family
    Size  = 77                          # em height in pixels
    Style = "Bold"                      # Regular | Bold | Italic | BoldItalic
    Y     = 70                          # vertical position (px from top)
}
```

Notes:

- `Size` is in **pixels**, not points — the template is a 192 DPI image,
  so point sizes would render twice as large as expected.
- `Name` is just an identifier for readability; `Text` is what appears.

---

## Participant fields — `$CertificateFields`

The label-above-value groups showing participant data. Two sets, one per
script:

| Script | Set | Fields |
|--------|-----|--------|
| `GenerateCertificates.ps1` | `$CertificateFields.Standard` | Name, Distance |
| `GenerateCertificatesTIME.ps1` | `$CertificateFields.Time` | Name, Distance, Time |

Each entry:

```powershell
@{
    Name       = "Distance"       # participant column the value is read from
    Label      = "Distance"       # text drawn above the value
    Suffix     = " metres"        # appended to the value (optional)
    LabelFont  = "Arial"
    LabelSize  = $FieldLabelSize
    LabelStyle = "Bold"
    LabelY     = 1265
    ValueFont  = "Arial"
    ValueSize  = $FieldValueSize
    ValueStyle = "Regular"
    ValueY     = 1380
}
```

| Key | Meaning |
|-----|---------|
| `Name` | Participant column read from the spreadsheet (`No`, `Name`, `Distance`, `Time`) |
| `Label` | The small text drawn above the value |
| `Suffix` | Optional text appended to the value (e.g. `" metres"`) |
| `LabelFont` / `LabelSize` / `LabelStyle` / `LabelY` | Label font, size, style and position |
| `ValueFont` / `ValueSize` / `ValueStyle` / `ValueY` | Value font, size, style and position |

Current positions:

| Field | Label Y | Value Y |
|-------|---------|---------|
| Name (standard) | 1050 | 1150 |
| Distance (standard) | 1265 | 1380 |
| Name (time) | 1030 | 1110 |
| Distance (time) | 1190 | 1270 |
| Time (time) | 1350 | 1430 |

---

## Field font sizes

All field labels and values share a single base size, with an optional
delta for the values:

```powershell
$LabelValueSize = 48        # base size applied to every label and value
$ValueSizeDelta = 0         # +/- adjustment applied to values only

$FieldLabelSize = $LabelValueSize
$FieldValueSize = $LabelValueSize + $ValueSizeDelta
```

| `$ValueSizeDelta` | Effect |
|-------------------|--------|
| `0` | Labels and values the same size (default) |
| `+n` | Values larger than labels |
| `-n` | Values smaller than labels |

The only constraint is that the result stays above zero — the font
constructor rejects a size of 0 or less.

To give labels and values completely independent sizes, assign numbers
directly:

```powershell
$FieldLabelSize = 44
$FieldValueSize = 52
```

---

## Adding a field

To draw another participant column (for example `No`), add an entry to
the appropriate set — the script draws whatever the set contains:

```powershell
@{
    Name       = "No"
    Label      = "Number"
    LabelFont  = "Arial"
    LabelSize  = $FieldLabelSize
    LabelStyle = "Bold"
    LabelY     = 1450
    ValueFont  = "Arial"
    ValueSize  = $FieldValueSize
    ValueStyle = "Regular"
    ValueY     = 1510
}
```

Fields are drawn in array order.
