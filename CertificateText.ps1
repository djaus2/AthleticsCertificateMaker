#
# Certificate header text definitions.
# Dot-sourced by GenerateCertificates.ps1 and GenerateCertificatesTIME.ps1.
#
# Each item is drawn centred horizontally at Y, at its natural width.
# To change the event wording, edit Text here once - both scripts pick it up.
#
$CertificateText = @(
    @{
        Name  = "Title"
        Text  = "1 HOUR TRACK RUN"
        Font  = "Times New Roman"
        Size  = 77
        Style = "Bold"
        Y     = 70
    }
    @{
        Name  = "SubTitle"
        Text  = "Moonee Valley Athletics Centre"
        Font  = "Arial Narrow"
        Size  = 51
        Style = "Regular"
        Y     = 180
    }
    @{
        Name  = "Host"
        Text  = "Aberfeldie Masters Athletics"
        Font  = "Times New Roman"
        Size  = 38
        Style = "Italic"
        Y     = 247
    }
    @{
        Name  = "Date"
        Text  = "16 September 2026"
        Font  = "Arial Narrow"
        Size  = 34
        Style = "Regular"
        Y     = 312
    }
)
