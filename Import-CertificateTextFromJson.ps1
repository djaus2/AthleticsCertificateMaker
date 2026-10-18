# Import variables and collections from a JSON file into the PowerShell session.
# Dot-source this file to make the variables available to the caller, e.g.
#   . .\Import-CertificateTextFromJson.ps1 -JsonPath .\CertificateText.json
param(
	[string]$JsonPath = "CertificateText.json"
)

if (-not (Test-Path -LiteralPath $JsonPath)) {
	throw "JSON file not found: $JsonPath"
}

$data = Get-Content -Raw -LiteralPath $JsonPath | ConvertFrom-Json -Depth 10

# Map top-level scalars into variables so existing generator scripts can use them directly.
$mappings = @(
	'CertificateTemplateFile','ParticipantCsvFile','ParticipantCsvDelimiter','SendUsingAccount',
	'DefaultTextColor','DefaultLabelColor','DefaultValueColor',
	'DefaultLabelTextSize','DefaultDataTextSize',
	'CertificateSmoothingMode','CertificateTextRenderingHint','ShowLocation','ShowHost'
)

foreach ($name in $mappings) {
	if ($data.PSObject.Properties.Name -contains $name) {
		Set-Variable -Name $name -Value $data.$name -Scope Global -Force
	}
}

# Collections
Set-Variable -Name CertificateText -Value $data.CertificateText -Scope Global -Force
Set-Variable -Name CertificateFields -Value $data.CertificateFields -Scope Global -Force

function Get-ParticipantCsvDelimiter {
	param([string]$Path)

	switch ($ParticipantCsvDelimiter.ToLowerInvariant()) {
		"comma" { return ',' }
		"tab" { return "`t" }
		"auto" {
			$HeaderLine = Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop
			if ([string]::IsNullOrWhiteSpace($HeaderLine)) {
				throw "Participant file is empty: $Path"
			}

			$CommaCount = ([regex]::Matches($HeaderLine, ',')).Count
			$TabCount = ([regex]::Matches($HeaderLine, "`t")).Count

			if ($CommaCount -eq $TabCount) {
				throw "Could not detect a comma or tab delimiter in the header of '$Path'."
			}
			if ($CommaCount -gt $TabCount) { return ',' }
			return "`t"
		}
		default {
			throw "Invalid ParticipantCsvDelimiter '$ParticipantCsvDelimiter'. Use Auto, Comma, or Tab."
		}
	}
}

function Import-ParticipantCsv {
	param([string]$Path)

	$Delimiter = Get-ParticipantCsvDelimiter -Path $Path
	Import-Csv -LiteralPath $Path -Delimiter $Delimiter
}

Write-Output "Imported variables: $($mappings -join ', '), CertificateText, CertificateFields (scope: Global)"
