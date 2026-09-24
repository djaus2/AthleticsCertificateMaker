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
# To size values differently from labels, set a +/- ValueSizeDelta.
#
$LabelValueSize = 48
$ValueSizeDelta = 0

$FieldLabelSize = $LabelValueSize
$FieldValueSize = $LabelValueSize + $ValueSizeDelta

$CertificateFields = @{
    Standard = @(
        @{
            Name       = "Name"
            Label      = "Name"
            LabelFont  = "Arial"
            LabelSize  = $FieldLabelSize
            LabelStyle = "Bold"
            LabelY     = 1050
            ValueFont  = "Arial"
            ValueSize  = $FieldValueSize
            ValueStyle = "Regular"
            ValueY     = 1150
        }
        @{
            Name       = "Distance"
            Label      = "Distance"
            Suffix     = " metres"
            LabelFont  = "Arial"
            LabelSize  = $FieldLabelSize
            LabelStyle = "Bold"
            LabelY     = 1265
            ValueFont  = "Arial"
            ValueSize  = $FieldValueSize
            ValueStyle = "Regular"
            ValueY     = 1380
        }
    )
    Time = @(
        @{
            Name       = "Name"
            Label      = "Name"
            LabelFont  = "Arial"
            LabelSize  = $FieldLabelSize
            LabelStyle = "Bold"
            LabelY     = 1030
            ValueFont  = "Arial"
            ValueSize  = $FieldValueSize
            ValueStyle = "Regular"
            ValueY     = 1110
        }
        @{
            Name       = "Distance"
            Label      = "Distance"
            Suffix     = " metres"
            LabelFont  = "Arial"
            LabelSize  = $FieldLabelSize
            LabelStyle = "Bold"
            LabelY     = 1190
            ValueFont  = "Arial"
            ValueSize  = $FieldValueSize
            ValueStyle = "Regular"
            ValueY     = 1270
        }
        @{
            Name       = "Time"
            Label      = "Time"
            LabelFont  = "Arial"
            LabelSize  = $FieldLabelSize
            LabelStyle = "Bold"
            LabelY     = 1350
            ValueFont  = "Arial"
            ValueSize  = $FieldValueSize
            ValueStyle = "Regular"
            ValueY     = 1430
        }
    )
}
