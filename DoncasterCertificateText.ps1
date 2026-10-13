#
# Shared certificate and event settings.
# Dot-sourced by the certificate generation and mail scripts.
#
# Each item is drawn centred horizontally at Y, at its natural width.
# To change the event wording, edit Text here once - both scripts pick it up.
#
#  Use CertificateTemplateBlank.png instead of the default CertificateTemplate.png 
#  ... if you want to create a new certificate for a new event and so add a new background image to that.
#
#$CertificateTemplateFile = "CertificateTemplateBlank.png"
$CertificateTemplateFile = "DoncasterCertificateTemplate2.png"

$ParticipantCsvFile = "results.csv"
$SendUsingAccount = "djaus@sportronics.com.au"

$DefaultTextColor = "Red"
$DefaultLabelColor = "LightGreen"
$DefaultValueColor = "Yellow"

$DefaultLabelTextSize = 70
$DefaultDataTextSize = 90

$CertificateSmoothingMode = "HighQuality"
$CertificateTextRenderingHint = "AntiAliasGridFit"

$CertificateText = @(
    @{
        Name  = "Title"
        Text  = "1 HOUR TRACK RUN"
        Color = $DefaultTextColor
        Font  = "Times New Roman"
        Size  = 90
        Style = "Bold"
        Y     = 70
    }
    @{
        Name  = "SubTitle"
        Text  = "Doncaster Athletics Centre"
        Color = "White"
        Font  = "Arial Narrow"
        Size  = 38
        Style = "Regular"
        Y     = 340
    }
    @{
        Name  = "Host"
        Text  = ""
        Color = $DefaultTextColor
        Font  = "Times New Roman"
        Size  = 38
        Style = "Italic"
        Y     = 360
    }
    @{
        Name  = "Date"
        Text  = "5th October 2026"
        Color = "LightGreen"
        Font  = "Arial Narrow"
        Size  = 51
        Style = "Regular"
        Y     = 410
    }
)

#
# Certificate field definitions - a label above each participant value.
# Each script uses its own set:
#   GenerateCertificates.ps1      -> $CertificateFields.Standard
#   GenerateCertificatesTIME.ps1  -> $CertificateFields.Time
#
# Name   = participant column the value is read from (No/Name/Distance/Time)
# Label  = text drawn above the value
# Suffix = appended to the value if present (e.g. " metres")
#
# Shared font size for all field labels and values - change once here.
#


$CertificateFields = @{
    Standard = @(
        @{
            Name       = "Name"
            Label      = "Name:"
            LabelFont  = "Arial"
            LabelSize  = $DefaultLabelTextSize
            LabelStyle = "Bold"
            LabelColor = $DefaultLabelColor
            LabelY     = 800
            ValueFont  = "Arial"
            ValueSize  = $DefaultDataTextSize
            ValueStyle = "Regular"
            ValueColor = $DefaultValueColor
            ValueY     = 900
        }
        @{
            Name       = "Distance"
            Label      = "Distance"
            Suffix     = " metres"
            LabelFont  = "Arial"
            LabelSize  = $DefaultLabelTextSize
            LabelStyle = "Bold"
            LabelColor = $DefaultLabelColor
            LabelY     = 1100
            ValueFont  = "Arial"
            ValueSize  = $DefaultDataTextSize
            ValueStyle = "Regular"
            ValueColor = $DefaultValueColor
            ValueY     = 1200
        }
    )
    Time = @(
        @{
            Name       = "Name"
            Label      = "Name"
            LabelFont  = "Arial"
            LabelSize  = $DefaultLabelTextSize
            LabelStyle = "Bold"
            LabelColor = $DefaultLabelColor
            LabelY     = 800
            ValueFont  = "Arial"
            ValueSize  = $DefaultDataTextSize
            ValueStyle = "Regular"
            ValueColor = $DefaultValueColor
            ValueY     = 900
        }
        @{
            Name       = "Distance"
            Label      = "Distance"
            Suffix     = " metres"
            LabelFont  = "Arial"
            LabelSize  = $DefaultLabelTextSize
            LabelStyle = "Bold"
            LabelColor = $DefaultLabelColor
            LabelY     = 1050
            ValueFont  = "Arial"
            ValueSize  = $DefaultDataTextSize
            ValueStyle = "Regular"
            ValueColor = $DefaultValueColor
            ValueY     = 1150
        }
        @{
            Name       = "Time"
            Label      = "Time"
            LabelFont  = "Arial"
            LabelSize  = $DefaultLabelTextSize
            LabelStyle = "Bold"
            LabelColor = $DefaultLabelColor
            LabelY     = 1300
            ValueFont  = "Arial"
            ValueSize  = $DefaultDataTextSize
            ValueStyle = "Regular"
            ValueColor = $DefaultValueColor
            ValueY     = 1400
        }
    )
}
