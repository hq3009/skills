<#
.SYNOPSIS
    Create symbolic links under the skills directories of CodeBuddy / workbuddy
    for every skill (a directory containing SKILL.md) found in this repo.

.DESCRIPTION
    Discovery is recursive: any directory containing SKILL.md is treated as a
    skill, and is never descended into (a skill directory is a leaf unit).
    Third-party skills maintained as submodules (e.g. under vendor-skills/) are
    discovered the same way, including collection repos that keep several
    skills under a skills/ subfolder.

    Collection repos may also keep helper directories (e.g. "_shared") that
    their skills reference via "../" relative paths. Underscore-prefixed
    sibling directories of a discovered skill that contain no SKILL.md are
    linked as well, so those relative references keep working.

    For each target skills directory:
      - links pointing into this repo that no longer match a discovered skill
        or helper directory are removed (stale links);
      - every discovered skill/helper gets a link named after its directory.
    Existing links are removed first (the real directory is untouched) and
    recreated.

.NOTES
    Creating symbolic links on Windows requires admin privileges by default,
    or Developer Mode enabled. If a symbolic link cannot be created, a
    directory junction is used as a fallback.
#>

[CmdletBinding()]
param(
    # Repo root, defaults to the script's own directory
    [string]$RepoRoot = $PSScriptRoot,

    # Target skills directories; defaults to CodeBuddy and workbuddy user dirs
    [string[]]$SkillsDirs = @(
        (Join-Path $env:USERPROFILE '.codebuddy\skills'),
        (Join-Path $env:USERPROFILE '.workbuddy\skills')
    )
)

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path $RepoRoot).Path.TrimEnd('\')

# Create a link at $linkPath pointing to $sourceDir, replacing any existing
# entry. Uses .NET delete APIs so that only the reparse point is removed.
function New-LinkItem {
    param([string]$linkPath, [string]$sourceDir)

    if (Test-Path $linkPath) {
        $existingItem = Get-Item $linkPath -Force
        if ($existingItem.LinkType) {
            if ($existingItem.PSIsContainer) {
                [System.IO.Directory]::Delete($linkPath)
            }
            else {
                [System.IO.File]::Delete($linkPath)
            }
        }
        else {
            Remove-Item $linkPath -Recurse -Force
        }
    }

    try {
        New-Item -ItemType SymbolicLink -Path $linkPath -Target $sourceDir | Out-Null
        Write-Host "Symbolic link created: $linkPath -> $sourceDir"
    }
    catch {
        New-Item -ItemType Junction -Path $linkPath -Target $sourceDir | Out-Null
        Write-Host "Junction created (symlink unavailable): $linkPath -> $sourceDir"
    }
}

# Discover skill directories: any directory containing SKILL.md.
$skipNames = @('.git', 'node_modules', '__pycache__', '.codebuddy', '.workbuddy')
$skillDirs = New-Object System.Collections.Generic.List[string]
$stack = [System.Collections.Generic.Stack[string]]::new()
$stack.Push($RepoRoot)

while ($stack.Count -gt 0) {
    $dir = $stack.Pop()
    if (Test-Path (Join-Path $dir 'SKILL.md')) {
        $skillDirs.Add($dir)
        continue   # skill directory is a leaf unit; do not scan inside it
    }
    Get-ChildItem -Path $dir -Directory -Force |
        Where-Object { $skipNames -notcontains $_.Name } |
        ForEach-Object { $stack.Push($_.FullName) }
}

if ($skillDirs.Count -eq 0) {
    Write-Warning "There's no skill directory containing SKILL.md under $RepoRoot."
    return
}

# Discover helper directories referenced via "../": underscore-prefixed
# siblings of a skill directory that contain no SKILL.md (e.g. "_shared").
$helperDirs = @()
foreach ($d in $skillDirs) {
    $parent = Split-Path $d -Parent
    $helperDirs += Get-ChildItem -Path $parent -Directory -Force |
        Where-Object {
            $_.Name.StartsWith('_') -and
            -not (Test-Path (Join-Path $_.FullName 'SKILL.md'))
        } |
        ForEach-Object { $_.FullName }
}
$helperDirs = @($helperDirs | Sort-Object -Unique)

$linkedNames = @{}
foreach ($d in ($skillDirs + $helperDirs)) {
    $linkedNames[[System.IO.Path]::GetFileName($d)] = $true
}

Write-Host "Discovered $($skillDirs.Count) skill(s):"
$skillDirs | ForEach-Object { Write-Host "  $_" }
if ($helperDirs.Count -gt 0) {
    Write-Host "Discovered $($helperDirs.Count) helper directorie(s):"
    $helperDirs | ForEach-Object { Write-Host "  $_" }
}

foreach ($targetRoot in $SkillsDirs) {
    if (-not (Test-Path $targetRoot)) {
        New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
    }

    # Remove stale links that point into this repo but match nothing discovered
    Get-ChildItem -Path $targetRoot -Force |
        Where-Object {
            $_.LinkType -and $_.Target -and
            @($_.Target | Where-Object { $_ -like "$RepoRoot*" }) -and
            -not $linkedNames.ContainsKey($_.Name)
        } |
        ForEach-Object {
            if ($_.PSIsContainer) {
                [System.IO.Directory]::Delete($_.FullName)
            }
            else {
                [System.IO.File]::Delete($_.FullName)
            }
            Write-Host "Stale link removed: $($_.FullName)"
        }

    foreach ($skillDir in $skillDirs) {
        New-LinkItem -linkPath (Join-Path $targetRoot ([System.IO.Path]::GetFileName($skillDir))) `
                     -sourceDir $skillDir
    }

    foreach ($helperDir in $helperDirs) {
        New-LinkItem -linkPath (Join-Path $targetRoot ([System.IO.Path]::GetFileName($helperDir))) `
                     -sourceDir $helperDir
    }
}
