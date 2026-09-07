<#
.SYNOPSIS
    Create symbolic links under CodeBuddy's skills directory for each skill (a directory containing SKILL.md) in the current repo.

.DESCRIPTION
    Scans the top-level subdirectories of the repo root for directories containing SKILL.md (i.e. skill directories),
    and creates a symbolic link under $env:USERPROFILE\.codebuddy\skills pointing to each skill's source directory.

.NOTES
    Creating symbolic links on Windows requires admin privileges by default, or Developer Mode enabled.
    If a link already exists, it is removed first (the real directory is untouched) and recreated.
#>

[CmdletBinding()]
param(
    # Repo root, defaults to the script's own directory
    [string]$RepoRoot = $PSScriptRoot,

    # CodeBuddy skills directory; defaults to the current user's profile path (username generalized)
    [string]$CodeBuddySkillsDir = (Join-Path $env:USERPROFILE '.codebuddy\skills')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $CodeBuddySkillsDir)) {
    New-Item -ItemType Directory -Path $CodeBuddySkillsDir -Force | Out-Null
}

# Find all top-level subdirectories under the repo that contain SKILL.md
$skillDirs = Get-ChildItem -Path $RepoRoot -Directory |
Where-Object { Test-Path (Join-Path $_.FullName 'SKILL.md') }

if (-not $skillDirs) {
    Write-Warning "There's no skill directory containing SKILL.md under $RepoRoot."
    return
}

foreach ($skillDir in $skillDirs) {
    $linkPath = Join-Path $CodeBuddySkillsDir $skillDir.Name

    if (Test-Path $linkPath) {
        $existingItem = Get-Item $linkPath -Force
        if ($existingItem.LinkType) {
            Remove-Item $linkPath -Force
        }
        else {
            Remove-Item $linkPath -Recurse -Force
        }
    }

    New-Item -ItemType SymbolicLink -Path $linkPath -Target $skillDir.FullName | Out-Null
    Write-Host "Symbolic link created: $linkPath -> $($skillDir.FullName)"
}
