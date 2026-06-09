# release.ps1 - wxd.files.winforms Auto Release Tool
# Usage: Double-click release.bat (recommended), or right-click -> Run with PowerShell
# Prerequisite: On first run, you will be prompted to enter your Gitee personal token, which will be saved automatically.

param(
    [string]$Version = "",
    [string]$ReleaseNotes = ""
)

# ============================================
# Configuration
# ============================================
$GiteeOwner = "wxd-gaming"              # Gitee username/organization
$GiteeRepo  = "wxd.files.winforms"      # Repository name
$GiteeToken = $env:GITEE_TOKEN          # Read token from environment variable

$ProjectDir  = Join-Path $PSScriptRoot "wxd.files.winforms"
$CsprojPath  = Join-Path $ProjectDir "wxd.files.winforms.csproj"
$PublishDir  = Join-Path $PSScriptRoot "publish"


# ============================================
# 0. Pre-checks
# ============================================
Write-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  wxd.files.winforms Auto Release Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host

# Check Gitee Token
if (-not $GiteeToken) {
    Write-Host "[INFO] GITEE_TOKEN environment variable not found." -ForegroundColor Yellow
    Write-Host "       A Gitee personal token is required to create a Release." -ForegroundColor Yellow
    Write-Host "       How to get one: Gitee website -> Settings -> Private Token -> Generate" -ForegroundColor Yellow
    Write-Host "       Required scope: projects" -ForegroundColor Yellow
    Write-Host

    $inputToken = Read-Host "Please enter your Gitee personal token"

    if (-not $inputToken) {
        Write-Host "[ERROR] Token cannot be empty. Cancelled." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        exit 1
    }

    # Save to user environment variable persistently
    try {
        [Environment]::SetEnvironmentVariable("GITEE_TOKEN", $inputToken, "User")
        $env:GITEE_TOKEN = $inputToken
        $GiteeToken = $inputToken
        Write-Host "[OK] Token saved to user environment variable. You won't need to enter it again." -ForegroundColor Green
    } catch {
        Write-Host "[WARN] Could not save token persistently. It will only be valid for this session." -ForegroundColor Yellow
        $env:GITEE_TOKEN = $inputToken
        $GiteeToken = $inputToken
    }
    Write-Host
}

# Check csproj file
if (-not (Test-Path $CsprojPath)) {
    Write-Host "[ERROR] Project file not found: $CsprojPath" -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

# ============================================
# 1. Input version number
# ============================================
if (-not $Version) {
    # Read current version from latest git tag as reference
    $currentTag = ""
    try {
        $currentTag = git describe --tags --abbrev=0 2>$null
        if ($currentTag -match "^v") {
            $currentTag = $currentTag.Substring(1)
        }
    } catch {}
    if ($currentTag) {
        Write-Host "Current version: $currentTag" -ForegroundColor Yellow
    }
    $Version = Read-Host "Enter new version number (e.g. 1.0.2.4)"
}

if (-not $Version) {
    Write-Host "[ERROR] Version number cannot be empty." -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

$TagName = "v$Version"
$ExeName = "wxd.files.winforms.exe"

# ============================================
# 2. Input release notes
# ============================================
if (-not $ReleaseNotes) {
    Write-Host
    Write-Host "Enter release notes (multi-line, empty line to finish):" -ForegroundColor Yellow
    $lines = @()
    while ($true) {
        $line = Read-Host ">"
        if (-not $line) { break }
        $lines += $line
    }
    $ReleaseNotes = $lines -join "`n"
}

if (-not $ReleaseNotes.Trim()) {
    Write-Host "[WARN] Release notes is empty, using default." -ForegroundColor Yellow
    $ReleaseNotes = "Release $TagName"
}

# ============================================
# 3. Confirm information
# ============================================
Write-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Release Information" -ForegroundColor Cyan
Write-Host "========================================"
Write-Host "  Version : $Version" -ForegroundColor White
Write-Host "  Tag     : $TagName" -ForegroundColor White
Write-Host "  File    : $ExeName" -ForegroundColor White
Write-Host "  Notes   :" -ForegroundColor White
foreach ($line in ($ReleaseNotes -split "`n")) {
    Write-Host "            $line" -ForegroundColor Gray
}
Write-Host "========================================"
Write-Host

$confirm = Read-Host "Confirm release? (y/n)"
if ($confirm -ne 'y' -and $confirm -ne 'Y') {
    Write-Host "Release cancelled." -ForegroundColor Yellow
    Read-Host "Press any key to exit..."
    exit 0
}

# ============================================
# 4. Build project
# ============================================
Write-Host
Write-Host "[1/6] Building project (dotnet publish)..." -ForegroundColor Cyan
Write-Host

# Clean old publish directory
if (Test-Path $PublishDir) {
    Remove-Item -Recurse -Force $PublishDir
}

$publishArgs = @(
    "publish", $CsprojPath,
    "-c", "Release",
    "-o", $PublishDir,
    "-p:Version=$Version",
    "--self-contained", "false",
    "-p:PublishSingleFile=true",
    "-p:IncludeNativeLibrariesForSelfExtract=true"
)

$publishOutput = & dotnet $publishArgs 2>&1
$publishExitCode = $LASTEXITCODE

# Output build result
if ($publishOutput) {
    foreach ($line in $publishOutput) {
        if ($line -match "error") {
            Write-Host $line -ForegroundColor Red
        } elseif ($line -match "warning") {
            Write-Host $line -ForegroundColor Yellow
        } else {
            Write-Host $line -ForegroundColor Gray
        }
    }
}

if ($publishExitCode -ne 0) {
    Write-Host "[ERROR] Build failed, exit code: $publishExitCode" -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

Write-Host "[1/6] Build successful!" -ForegroundColor Green

# ============================================
# 5. Locate published EXE (single-file mode)
# ============================================
Write-Host
Write-Host "[2/6] Locating published EXE..." -ForegroundColor Cyan

$ExeName = "wxd.files.winforms.exe"
$ExePath = Join-Path $PublishDir $ExeName

if (-not (Test-Path $ExePath)) {
    Write-Host "[ERROR] Published EXE not found: $ExePath" -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

$exeSize = [math]::Round((Get-Item $ExePath).Length / 1MB, 2)
Write-Host "[2/6] EXE found: $ExeName ($exeSize MB)" -ForegroundColor Green

# ============================================
# 6. Generate MD5 (of EXE, not ZIP)
# ============================================
Write-Host
Write-Host "[3/6] Generating MD5 of EXE..." -ForegroundColor Cyan

$md5 = (Get-FileHash $ExePath -Algorithm MD5).Hash.ToLower()

Write-Host "[3/6] MD5 generated: $md5" -ForegroundColor Green

# ============================================
# 7. Create and push Git Tag (no md5.txt commit)
# ============================================
Write-Host
Write-Host "[4/6] Creating Git Tag..." -ForegroundColor Cyan

# Create tag only (no commit of md5.txt)
git tag -a $TagName -m "$ReleaseNotes" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[WARN] Tag creation may have failed (may already exist), continuing..." -ForegroundColor Yellow
}

# Push tag only
Write-Host "  Pushing tag to remote..."
git push origin $TagName 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

Write-Host "[4/6] Tag pushed" -ForegroundColor Green

# ============================================
# 8. Create Gitee Release via API
# ============================================
Write-Host
Write-Host "[5/6] Creating Gitee Release..." -ForegroundColor Cyan

$releaseBody = @"
## version $TagName

$ReleaseNotes

---

**MD5**: `$md5`
**File**: $ExeName
**Size**: $exeSize MB
**Time**: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@

$apiBody = @{
    access_token     = $GiteeToken
    tag_name         = $TagName
    name             = $TagName
    body             = $releaseBody
    target_commitish = "master"
    prerelease       = $false
} | ConvertTo-Json -Compress

try {
    $releaseResponse = Invoke-RestMethod `
        -Uri "https://gitee.com/api/v5/repos/$GiteeOwner/$GiteeRepo/releases" `
        -Method Post `
        -Body $apiBody `
        -ContentType "application/json; charset=utf-8"

    $releaseId = $releaseResponse.id
    # Fallback: if html_url is empty, construct it manually
    $releaseUrl = if ($releaseResponse.html_url) { 
        $releaseResponse.html_url 
    } else { 
        "https://gitee.com/$GiteeOwner/$GiteeRepo/releases/$TagName" 
    }
    Write-Host "[5/6] Release created successfully!" -ForegroundColor Green
    Write-Host "  Release ID : $releaseId" -ForegroundColor Gray
    Write-Host "  URL        : $releaseUrl" -ForegroundColor Cyan
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host "[ERROR] Create Release failed: $errorMsg" -ForegroundColor Red

    if ($errorMsg -match "401|403|Unauthorized") {
        Write-Host "  Please check that GITEE_TOKEN is correct and has 'projects' permission." -ForegroundColor Yellow
    }
    if ($errorMsg -match "422|already|exist") {
        Write-Host "  A Release for this Tag may already exist. Please check." -ForegroundColor Yellow
    }

    Read-Host "Press any key to exit..."
    exit 1
}

# ============================================
# 9. Upload EXE as attachment (Win10/Win11 compatible)
# ============================================
Write-Host
Write-Host "[6/6] Uploading EXE as attachment..." -ForegroundColor Cyan

try {
    # Use curl.exe (built-in on Win10 1803+ / Win11) for reliable multipart upload
    $uploadUrl = "https://gitee.com/api/v5/repos/$GiteeOwner/$GiteeRepo/releases/$releaseId/attach_files"

    $responseFile = Join-Path $env:TEMP "gitee_upload_response.txt"
    $httpCodeFile = Join-Path $env:TEMP "gitee_upload_httpcode.txt"

    $httpCode = & curl.exe -s -w "%{http_code}" -X POST $uploadUrl `
        -F "access_token=$GiteeToken" `
        -F "file=@$ExePath;filename=$ExeName" `
        -o $responseFile 2>&1

    $responseBody = ""
    if (Test-Path $responseFile) {
        $responseBody = Get-Content $responseFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
        Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
    }

    if ($httpCode -eq "201" -or $httpCode -eq "200") {
        Write-Host "[6/6] EXE uploaded successfully!" -ForegroundColor Green
    } else {
        Write-Host "[WARN] EXE upload failed (HTTP $httpCode)" -ForegroundColor Yellow
        Write-Host "  Response: $responseBody" -ForegroundColor Gray
        Write-Host "  You can upload $ExeName manually on the Release page later." -ForegroundColor Yellow
    }
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host "[WARN] EXE upload failed: $errorMsg" -ForegroundColor Yellow
    Write-Host "  You can upload $ExeName manually on the Release page later." -ForegroundColor Yellow
}

# ============================================
# Done
# ============================================
Write-Host
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Release Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Version : $TagName" -ForegroundColor White
Write-Host "  MD5     : $md5" -ForegroundColor White
Write-Host "  File    : $ExeName ($exeSize MB)" -ForegroundColor White
Write-Host "  Release : $releaseUrl" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Green
Write-Host

Read-Host "Press any key to exit..."
