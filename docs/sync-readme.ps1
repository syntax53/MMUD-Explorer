# Rebuilds the version history in README.md from docs/changelog.txt (the source of truth).
#
#   powershell -ExecutionPolicy Bypass -File docs\sync-readme.ps1
#
# README keeps its own intro (everything above its first "vX.Y" heading); everything from
# there down is replaced with the changelog, with the changelog's long dash rules shortened
# to README's 42-dash rule. README is written UTF-8 (no BOM) with LF endings (.gitattributes).
# Trailing double spaces are kept on purpose: they are markdown line breaks on GitHub.

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$readmePath = Join-Path $root 'README.md'
$changelogPath = Join-Path $root 'docs\changelog.txt'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$readme = [IO.File]::ReadAllText($readmePath, $utf8) -split "\r?\n"
$firstVersion = -1
for ($i = 0; $i -lt $readme.Count; $i++) {
    if ($readme[$i] -match '^v\d') { $firstVersion = $i; break }
}
if ($firstVersion -lt 1) { throw "Could not find the first version heading (vX.Y) in README.md" }

$changelog = [IO.File]::ReadAllText($changelogPath, $utf8) -split "\r?\n" | ForEach-Object {
    if ($_ -match '^-{20,}\s*$') { ('-' * 42) + '  ' } else { $_ }
}

$result = ($readme[0..($firstVersion - 1)] + $changelog) -join "`n"
$before = [IO.File]::ReadAllText($readmePath, $utf8)
if ($result -ceq $before) {
    "README.md already matches docs/changelog.txt"
} else {
    [IO.File]::WriteAllText($readmePath, $result, $utf8)
    "README.md updated from docs/changelog.txt"
}
