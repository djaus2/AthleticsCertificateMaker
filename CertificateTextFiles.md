# CertificateText JSON and helper scripts

This document describes the three new files added to the repository that provide a JSON-based source of truth for certificate text and helper scripts to consume or regenerate the original PowerShell representation.

## Files

- `CertificateText.json`
  - JSON representation of the certificate configuration that was previously authored in `CertificateText.ps1`.
  - Contains top-level scalars such as `CertificateTemplateFile`, `ParticipantCsvFile`, colour/size defaults and flags (`ShowLocation`, `ShowHost`).
  - Contains `CertificateText` (array of text items: Title, Location, Host, Date) and `CertificateFields` (object with `Standard` and `Time` arrays describing label/value fields).
  - Edit this file to change wording, fonts, colours, positions and field layout.

- `Prev/Generate-CertificateTextFromJson.ps1`
  - Generates a PowerShell file (default: `CertificateText.ps1`) from `CertificateText.json`.
  - Serialises scalars and collections into PowerShell literal syntax and emits the CSV helper functions used by the generator scripts (`Get-ParticipantCsvDelimiter`, `Import-ParticipantCsv`).
  - Usage:
	```powershell
	Set-Location .\Prev
	.\Generate-CertificateTextFromJson.ps1 -JsonPath .\CertificateText.json -OutputPath .\CertificateText.ps1
	```
  - Use this when you want a generated PS1 file for backward compatibility with scripts that dot-source `CertificateText.ps1`.
  - From `Prev`, running `.\Generate-CertificateTextFromJson.ps1` uses the local JSON copy and writes `CertificateText.ps1` in that folder. Strings are emitted as literal values, preserving quotes, dollar signs and backticks.

- `Import-CertificateTextFromJson.ps1`
  - Loads `CertificateText.json` directly into the current PowerShell session so existing generator scripts can reference the variables without creating a PS1 file.
  - Sets variables for the common scalars and collections (`CertificateText`, `CertificateFields`). The script currently assigns variables to the `Global` scope to maintain compatibility with scripts that expect module-level variables.
  - Recommended usage (dot-source to import into the caller session):
	```powershell
	. .\Import-CertificateTextFromJson.ps1 -JsonPath .\CertificateText.json
	```
  - If you prefer not to use `Global` scope, the script can be adjusted to only work when dot-sourced (variables set in caller scope). Ask if you want that change.

## How to use with existing generator scripts

- Replace an existing dot-source line like:
  ```powershell
  . .\CertificateText.ps1
  ```
  with the importer (recommended):
  ```powershell
  . .\Import-CertificateTextFromJson.ps1 -JsonPath .\CertificateText.json
  ```

- Or keep the current flow by generating a PS1 first and then dot-sourcing it:
  ```powershell
  .\Prev\Generate-CertificateTextFromJson.ps1 -JsonPath .\CertificateText.json -OutputPath .\CertificateText.ps1
  . .\CertificateText.ps1
  ```

## Notes and suggestions

- Editing `CertificateText.json` is the preferred way to update certificate text and layout. Either regenerate the PS1 or dot-source the importer afterwards.
- The importer currently sets variables in `Global` scope for compatibility. If you want the importer to avoid global variables and rely on dot-sourcing only, request that change and it will be implemented.
- If you want me to update `GenerateCertificates.ps1` to automatically dot-source the importer, I can make that change.
