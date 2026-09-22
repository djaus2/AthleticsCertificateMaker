#
# Certificate header text definitions.
# Dot-sourced by GenerateCertificates.ps1 and GenerateCertificatesTIME.ps1.
#
# Each item is drawn centred horizontally at Y and squeezed horizontally
# to Width, matching the text that was baked into the original template
# image. To change the event wording, edit Text here once - both scripts
# pick it up.
#
$CertificateText = @(
    @{
        Name  = "Title"
        Text  = "1 HOUR TRACK RUN"
        Font  = "Times New Roman"
        Size  = 100
        Style = "Bold"
        Y     = 70
        Width = 818
    }
    @{
        Name  = "SubTitle"
        Text  = "Moonee Valley Athletics Centre"
        Font  = "Arial Narrow"
        Size  = 46
        Style = "Regular"
        Y     = 180
        Width = 614
    }
    @{
        Name  = "Host"
        Text  = "Aberfeldie Masters Athletics"
        Font  = "Times New Roman"
        Size  = 38
        Style = "Italic"
        Y     = 247
        Width = 459
    }
    @{
        Name  = "Date"
        Text  = "16 September 2026"
        Font  = "Arial Narrow"
        Size  = 32
        Style = "Regular"
        Y     = 312
        Width = 266
    }
)
