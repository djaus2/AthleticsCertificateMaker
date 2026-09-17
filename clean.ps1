Param(
	[switch]$WhatIf
)

<#
Deletes all files and folders that Git reports as "ignored" in this repository.

Usage:
  .\clean.ps1            # actually delete files
  .\clean.ps1 -WhatIf   # show what would be deleted

This script runs from the repository root (where the script lives) and uses:
  git ls-files --others --ignored --exclude-standard

The script is defensive: it only deletes paths under the repository root
and will skip missing paths.
#>

Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
Push-Location $RepoRoot
try {
	if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
		Write-Error "git not found in PATH. Please run this script where git is available."
		exit 2
	}

	$raw = & git ls-files --others --ignored --exclude-standard -z 2>$null

	if ([string]::IsNullOrEmpty($raw)) {
		Write-Host "No ignored files found." -ForegroundColor Green
		return 0
	}

	$items = $raw -split "`0" | Where-Object { $_ -ne '' }

	foreach ($item in $items) {
		# Skip any items inside the .vs folder
		if ($item -match '^(?:\.vs)([\\/]|$)') {
			continue
		}
		# Build full path and resolve
		$full = Join-Path $RepoRoot $item
		# Try to resolve path; if it can't be resolved skip quietly
		$resolved = Resolve-Path -LiteralPath $full -ErrorAction SilentlyContinue
		if (-not $resolved) {
			continue
		}

		$fullPath = $resolved.Path

		# Safety: ensure the path is inside the repo root
		if (-not $fullPath.StartsWith($RepoRoot, [System.StringComparison]::InvariantCultureIgnoreCase)) {
			continue
		}

		if ($WhatIf) {
			Write-Host "Would remove: $fullPath"
			continue
		}

		# Attempt to get item; if missing or inaccessible, skip quietly
		$itemInfo = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
		if (-not $itemInfo) {
			continue
		}

		# Remove directory or file
		if ($itemInfo.PSIsContainer) {
			Remove-Item -LiteralPath $fullPath -Recurse -Force -ErrorAction SilentlyContinue
		} else {
			Remove-Item -LiteralPath $fullPath -Force -ErrorAction SilentlyContinue
		}
	}

	Write-Host "Completed." -ForegroundColor Green
} finally {
	Pop-Location
}
