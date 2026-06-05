# wxd.files.winforms - Generate md5.txt (version + MD5 + changelog)
# Usage: Right-click -> Run with PowerShell

$md5Path = Join-Path $PSScriptRoot 'md5.txt'
$exePath = Join-Path $PSScriptRoot 'wxd.files.winforms.exe'

Write-Host
Write-Host "Current directory: $PSScriptRoot"
Write-Host

if (-not (Test-Path $exePath)) {
    Write-Host "[ERROR] wxd.files.winforms.exe not found." -ForegroundColor Red
    Read-Host "Press any key to continue..."
    exit 1
}

# Read old version and last SHA
$oldVersion = ''
$lastSha = ''
if (Test-Path $md5Path) {
    $allLines = Get-Content $md5Path -Encoding UTF8
    $oldVersion = $allLines[0].Trim()
    if ($allLines.Count -ge 3) {
        $lastSha = $allLines[2].Trim()
    }
}
if ($oldVersion) {
    Write-Host "Current version: $oldVersion"
}

# Input new version
$newVersion = Read-Host "Enter new version (e.g. 1.0.0.5)"
if (-not $newVersion) {
    Write-Host "[ERROR] Version cannot be empty." -ForegroundColor Red
    Read-Host "Press any key to continue..."
    exit 1
}

# Compute MD5
Write-Host
Write-Host "Computing MD5..."
Write-Host
$md5 = (Get-FileHash $exePath -Algorithm MD5).Hash.ToLower()

# Fetch git commit log
Write-Host "Fetching git commit log..."
$ErrorActionPreference = 'SilentlyContinue'

# Build git log args
if ($lastSha) {
    $gitArgs = @('log', "$($lastSha)..HEAD")
}
else {
    $gitArgs = @('log', '-5')
}
$gitArgs += '--pretty=format:%H%x00%s', '--no-merges'

# Capture git output as raw bytes, then decode as UTF-8
$rawBytes = & {
    $ErrorActionPreference = 'SilentlyContinue'
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    & git $gitArgs 2>&1 | Out-String | ForEach-Object { $_ }
}

# Fallback: use raw git output directly
$raw = & {
    $prevEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $result = & git $gitArgs
    [Console]::OutputEncoding = $prevEncoding
    $result
}

# Parse results
$newSha = ''
$notes = @()
if ($raw) {
    if ($raw -is [array]) {
        $lines = $raw
    }
    else {
        $lines = $raw.Split("`n")
    }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if ($line) {
            # SHA is first 40 chars, message is the rest (after null separator from %x00)
            $nullIdx = $line.IndexOf([char]0)
            if ($nullIdx -ge 0) {
                $sha = $line.Substring(0, $nullIdx).Trim()
                $msg = $line.Substring($nullIdx + 1).Trim()
            }
            elseif ($line.Length -ge 40) {
                $sha = $line.Substring(0, 40).Trim()
                $msg = $line.Substring(40).Trim()
            }
            else {
                $sha = $line.Trim()
                $msg = ''
            }

            # First SHA as latest (for next incremental read)
            if ($i -eq 0 -and $sha) {
                $newSha = $sha
            }

            if ($msg) {
                $notes += '-' + $msg
            }
        }
    }
}

# Write md5.txt (UTF-8 without BOM)
$outputLines = @($newVersion, $md5)
if ($newSha) {
    $outputLines += $newSha
}
if ($notes) {
    $outputLines += $notes
}

$utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes(($outputLines -join "`r`n"))
[System.IO.File]::WriteAllBytes($md5Path, $utf8Bytes)

# Output results
Write-Host
Write-Host "========================="
Write-Host " md5.txt generated!"
Write-Host " Version: $newVersion"
Write-Host " MD5    : $md5"
if ($newSha) {
    Write-Host " SHA    : $newSha"
}
Write-Host "========================="
Write-Host
Write-Host "Release Notes:"
Write-Host "--------"
if ($notes) {
    foreach ($n in $notes) {
        Write-Host $n
    }
}
else {
    Write-Host "(no new commits)"
}
Write-Host "--------"
Write-Host
Read-Host "Press any key to continue..."
