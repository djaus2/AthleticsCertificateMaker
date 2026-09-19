Param(
	[switch]$WhatIf
)

<#
Deletes all generated content in this repository.
Generated folders are emptied but left in place:
  - Output\*         generated certificate images/PDFs
                     (keeps the tracked sample PNG)
  - Results\*        generated results workbook and PDF
  - logs\*           generated logs
  - ~$*.xlsx         Excel lock files
  - ,\               stray folder left by a buggy run (removed entirely)

Usage:
  .\clean.ps1            # actually delete files
  .\clean.ps1 -WhatIf    # show what would be deleted

The script is defensive: it only deletes paths under the repository root
and skips missing paths.
#>

Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Folders to empty: all files inside are deleted, folders are left in place
$ClearFolders = @(
	"Output",
	"Results",
	"logs"
)

# Paths deleted entirely, including the folder itself
$DeletePaths = @(
	","
)

# Root file patterns to delete (wildcards allowed)
$DeleteGlobs = @(
	"~`$*.xlsx"
)

# Names to keep even when matched by the rules above
$KeepNames = @(
	"004-Fred Nurk.png"
)

function Remove-RepoItem {
	param(
		[string]$FullPath,
		[switch]$Recurse
	)
	$resolved = Resolve-Path -LiteralPath $FullPath -ErrorAction SilentlyContinue
	if (-not $resolved) {
		return
	}
	$path = $resolved.Path

	# Safety: only delete inside the repo root
	if (-not $path.StartsWith($RepoRoot, [System.StringComparison]::InvariantCultureIgnoreCase)) {
		return
	}

	if ($WhatIf) {
		Write-Host "Would remove: $path"
		return
	}

	if ($Recurse) {
		Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
	} else {
		Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
	}
}

Push-Location $RepoRoot
try {
	# Delete every file inside the generated-content folders, leave folders
	foreach ($folder in $ClearFolders) {
		$folderPath = Join-Path $RepoRoot $folder
		if (-not (Test-Path -LiteralPath $folderPath)) {
			continue
		}
		Get-ChildItem -LiteralPath $folderPath -Recurse -Force -File |
			Where-Object { $KeepNames -notcontains $_.Name } |
			ForEach-Object { Remove-RepoItem -FullPath $_.FullName }
	}

	foreach ($path in $DeletePaths) {
		$fullPath = Join-Path $RepoRoot $path
		if (-not (Test-Path -LiteralPath $fullPath)) {
			continue
		}
		if ($WhatIf) {
			Get-ChildItem -LiteralPath $fullPath -Recurse -Force |
				ForEach-Object { Write-Host "Would remove: $($_.FullName)" }
		}
		Remove-RepoItem -FullPath $fullPath -Recurse
	}

	foreach ($glob in $DeleteGlobs) {
		$pattern = Join-Path $RepoRoot $glob
		Get-ChildItem -Path $pattern -Force -File -ErrorAction SilentlyContinue |
			Where-Object { $KeepNames -notcontains $_.Name } |
			ForEach-Object { Remove-RepoItem -FullPath $_.FullName }
	}

	Write-Host "Completed." -ForegroundColor Green
} finally {
	Pop-Location
}
