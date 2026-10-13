# Doncaster Text Size, Color and Locations

This document records the changes made to certificate typography, colors, positions (Y coordinates), and layout configurations in [DoncasterCertificateText.ps1](https://github.com/djaus2/AthleticsCertificateMaker/DoncasterCertificateText.ps1) _(not yet avi;able here)_ for the Doncaster event certificates over 2026-10-12 and 2026-10-13.

---

## 1. Global / Default Color & Size Settings

| Setting | Previous Value | Doncaster Value | Description |
|---|---|---|---|
| `$CertificateTemplateFile` | `"CertificateTemplate.png"` | `"DoncasterCertificateTemplate2.png"` | Template background image |
| `$ParticipantCsvFile` | `"results.xlsx"` (as spreadsheet) | `"results.csv"` | Data source moved from Excel to CSV |
| `$SendUsingAccount` | `"account@location.com.au"` | `"djaus@sportronics.com.au"` | Outlook sender account |
| `$DefaultTextColor` | `@(0, 20, 90)` (RGB) | `"Red"` | Base default text color |
| `$DefaultLabelColor` | *None* (used shared RGB) | `"LightGreen"` | Default color for participant field labels |
| `$DefaultValueColor` | *None* (used shared RGB) | `"Yellow"` | Default color for participant field values |
| `$DefaultLabelTextSize` | `48` (via `$LabelValueSize`) | `70` | Default font size (px) for field labels |
| `$DefaultDataTextSize` | `48` (via `$LabelValueSize`) | `90` | Default font size (px) for participant data values |

---

## 2. Header Text (`$CertificateText`)

All header items are centred horizontally at their specified vertical `Y` coordinate (in pixels from the top of the canvas).

| Field Name | Text | Font | Size (px) | Style | Color | Y Position | Notes / Changes |
|---|---|---|---|---|---|---|---|
| **Title** | `1 HOUR TRACK RUN` | Times New Roman | 90 | Bold | `$DefaultTextColor` (`"Red"`) | `70` | Size increased from 77 to 90 px; explicit color added |
| **SubTitle** | `Doncaster Athletics Centre` | Arial Narrow | 38 | Regular | `"White"` | `340` | Text changed from *"Moonee Valley Athletics Centre"*; size adjusted from 51 to 38 px; Y moved from 180 to 340; color set to `"White"` |
| **Host** | `""` | Times New Roman | 38 | Italic | `$DefaultTextColor` (`"Red"`) | `360` | Text cleared (previously *"Aberfeldie Masters Athletics"*); Y moved from 247 to 360 |
| **Date** | `5th October 2026` | Arial Narrow | 51 | Regular | `"LightGreen"` | `410` | Date changed from *"16 September 2026"*; size adjusted from 34 to 51 px; Y moved from 312 to 410; color set to `"LightGreen"` |

---

## 3. Participant Field Definitions (`$CertificateFields`)

Participant fields render a label above the participant's data value.

### 3.1 Standard Fields (`$CertificateFields.Standard`)
Used by [GenerateCertificates.ps1](GenerateCertificates.ps1).

| Field | Element | Text / Value Source | Font | Size (px) | Style | Color | Y Position |
|---|---|---|---|---|---|---|---|
| **Name** | Label | `"Name:"` | Arial | 70 (`$DefaultLabelTextSize`) | Bold | `$DefaultLabelColor` (`"LightGreen"`) | `800` (was 1050) |
| | Value | `$Participant.Name` | Arial | 90 (`$DefaultDataTextSize`) | Regular | `$DefaultValueColor` (`"Yellow"`) | `900` (was 1150) |
| **Distance** | Label | `"Distance"` | Arial | 70 (`$DefaultLabelTextSize`) | Bold | `$DefaultLabelColor` (`"LightGreen"`) | `1100` (was 1265) |
| | Value | `"$($Participant.Distance) metres"` | Arial | 90 (`$DefaultDataTextSize`) | Regular | `$DefaultValueColor` (`"Yellow"`) | `1200` (was 1380) |

### 3.2 Time Fields (`$CertificateFields.Time`)
Used by [GenerateCertificatesTIME.ps1](GenerateCertificatesTIME.ps1).

| Field | Element | Text / Value Source | Font | Size (px) | Style | Color | Y Position |
|---|---|---|---|---|---|---|---|
| **Name** | Label | `"Name"` | Arial | 70 (`$DefaultLabelTextSize`) | Bold | `$DefaultLabelColor` (`"LightGreen"`) | `800` (was 1030) |
| | Value | `$Participant.Name` | Arial | 90 (`$DefaultDataTextSize`) | Regular | `$DefaultValueColor` (`"Yellow"`) | `900` (was 1110) |
| **Distance** | Label | `"Distance"` | Arial | 70 (`$DefaultLabelTextSize`) | Bold | `$DefaultLabelColor` (`"LightGreen"`) | `1050` (was 1190) |
| | Value | `"$($Participant.Distance) metres"` | Arial | 90 (`$DefaultDataTextSize`) | Regular | `$DefaultValueColor` (`"Yellow"`) | `1150` (was 1270) |
| **Time** | Label | `"Time"` | Arial | 70 (`$DefaultLabelTextSize`) | Bold | `$DefaultLabelColor` (`"LightGreen"`) | `1300` (was 1350) |
| | Value | `$Participant.Time` | Arial | 90 (`$DefaultDataTextSize`) | Regular | `$DefaultValueColor` (`"Yellow"`) | `1400` (was 1430) |

---

## 4. Script & Workflow Updates Summary

1. **Excel Removal**:
   - Replaced dependency on `ImportExcel` / `Excel.Application` with native `Import-Csv` from [results.csv](results.csv).
   - Generated results PDF using headless Microsoft Edge via `Export-ResultsPdf.ps1`.
2. **Color Customization**:
   - Added support for named drawing colors via `[System.Drawing.Color]::FromName(...)`.
   - Introduced `$DefaultTextColor`, `$DefaultLabelColor`, and `$DefaultValueColor`.
   - Enabled per-element `Color` overrides on header items, labels, and participant values.
3. **Size Configuration**:
   - Replaced fixed `$LabelValueSize` and delta calculation with explicit `$DefaultLabelTextSize` and `$DefaultDataTextSize`.
4. **Script Parametrization**:
   - Added `$CertificateTextDefinitionFile = "CertificateText.ps1"` definition at the top of [GenerateCertificates.ps1](GenerateCertificates.ps1).
