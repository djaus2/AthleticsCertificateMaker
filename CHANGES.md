# Changes — 19 September 2026

Summary of today's session on the Aberfeldie Masters Athletics
certificate generator.

## Repository

- Work moved entirely to `C:\temp\AthleticsCertificateMaker`, a clone of
  <https://github.com/djaus2/AthleticsCertificateMaker>. The older
  `C:\Certificates` folder points at the old repo
  (`djaus2/AthsCertificateMaker`) and is no longer being modified.

## Certificate text moved from image to scripts

All text that was previously baked into the certificate template image
is now drawn by the scripts. `CertificateTemplate.png` is now pure
artwork (frame, singlet, logo) and `CertificateTemplateorig.png` is kept
untouched as the unmodified master.

| Name | Text | Font | Size | Y | Width |
|------|------|------|------|---|-------|
| Title | 1 HOUR TRACK RUN | Times New Roman Bold | 100 px | 70 | 818 |
| SubTitle | Moonee Valley Athletics Centre | Arial Narrow | 46 px | 180 | 614 |
| Host | Aberfeldie Masters Athletics | Times New Roman Italic | 38 px | 247 | 459 |
| Date | 16 September 2026 | Arial Narrow | 32 px | 312 | 266 |

These definitions live in `CertificateText.ps1`, which both
`GenerateCertificates.ps1` and `GenerateCertificatesTIME.ps1`
dot-source — so the event wording only needs editing in one place. Each
line is drawn centred and squeezed horizontally to the recorded width.

The same was done earlier for the field labels: the standard script
draws *Name*/*Distance* labels above their values, and the TIME script
draws *Name*/*Distance*/*Time* label-and-value groups.

## Matching the baked fonts

A PNG stores only pixels — there is no font metadata to recover, so the
fonts had to be identified by measurement and eyeball:

1. **Measure the baked glyphs.** Each text line's bounding box was found
   by scanning the image for the navy text colour, giving the exact
   pixel width, height and centre of every line.
2. **Compare candidates by width.** `Graphics.MeasureString()` was run
   over likely fonts (Times New Roman, Georgia, Cambria, Garamond,
   Arial Narrow, etc.) at a known size; the measured width was then
   scaled to the baked width to estimate the real font size.
3. **Compare letterforms visually.** Candidate fonts were rendered
   side-by-side with the original strip and matched by shape — this is
   how Arial Narrow (subtitle, date) and Times New Roman italic (host
   line) were identified.
4. **Reproduce the condensation.** The baked text is narrower than the
   same font at the same height would naturally render — it was
   horizontally compressed when the artwork was created. The scripts
   replicate this with `Graphics.ScaleTransform`, squeezing each line to
   its measured width while keeping the correct height.
5. **DPI pitfall.** The template is 192 DPI, so a point-based font
   renders twice as large as expected. All drawn text uses
   `GraphicsUnit.Pixel` to avoid the unit conversion entirely.

The result is visually near-identical to the baked version, but the
wording, fonts and positions are now all script-controlled — which is
what makes the package reusable for a different event.

## Other changes

- **Startup email-mode check** — each script scans its own source to
  detect whether `$Mail.Save()` and/or `$Mail.Send()` are enabled:
  both enabled aborts with an error popup; send-only requires
  confirmation; save-only proceeds silently; neither warns that no
  emails will be produced.
- **Sending account resolution** — `$SendUsingAccount` is resolved once
  at startup; if the account isn't in the Outlook profile a warning is
  shown and the default account is used, instead of silently assigning
  null. Applied at every mail site, including the certificate branch of
  `GenerateCertificates.ps1` which previously never set it.
- **`clean.ps1` rework** — empties `Output\*`, `Results\*`, `logs\*`,
  removes `~$*.xlsx` lock files and the stray `,\` folder, while keeping
  the tracked sample `004-Fred Nurk.png` and leaving the folder
  structure in place. `-WhatIf` previews without deleting.
- **Outlook requirements documented** — classic desktop Outlook with a
  default mail profile is required; the "New Outlook" app does not
  support COM automation.
