# release.ps1 - wxd.files.winforms Auto Release Tool (dual publish: Gitee + GitHub)
# Usage: Double-click release.bat (recommended), or right-click -> Run with PowerShell
# Prerequisite: On first run, you will be prompted for your Gitee and GitHub personal
# tokens (GITEE_TOKEN / GITHUB_TOKEN); both are saved to user environment variables.
# Leaving the GitHub token empty publishes to Gitee only.

param(
    [string]$Version = "",
    [string]$ReleaseNotes = ""
)

# ============================================
# Configuration
# ============================================
$GiteeOwner = "wxd-gaming"              # Gitee username/organization
$GiteeRepo  = "wxd.files.winforms"      # Repository name
$GiteeRemoteName = "origin"             # Gitee git remote

$GithubOwner = "wxd-gaming"             # GitHub username/organization
$GithubRepo  = "wxd.files.winforms"     # Repository name
$GithubRepoUrl = "https://github.com/$GithubOwner/$GithubRepo.git"

$ProjectDir  = Join-Path $PSScriptRoot "wxd.files.winforms"
$CsprojPath  = Join-Path $ProjectDir "wxd.files.winforms.csproj"
$PublishDir  = Join-Path $PSScriptRoot "publish"


# ============================================
# Helpers
# ============================================

# Read token from process/user env; prompt and persist when missing.
# Returns $null when the user skips it (caller decides whether that is fatal).
function Resolve-ApiToken {
    param(
        [string]$EnvName,
        [string]$DisplayName,
        [string]$ScopeHint
    )

    $token = [Environment]::GetEnvironmentVariable($EnvName, "Process")
    if (-not $token) { $token = [Environment]::GetEnvironmentVariable($EnvName, "User") }
    if ($token) { return $token }

    Write-Host "[INFO] $EnvName environment variable not found." -ForegroundColor Yellow
    Write-Host "       A $DisplayName personal token is required to create a Release." -ForegroundColor Yellow
    Write-Host "       How to get one: $DisplayName website -> Settings -> Developer settings -> Personal Access Token" -ForegroundColor Yellow
    Write-Host "       Required scope: $ScopeHint" -ForegroundColor Yellow
    Write-Host
    Write-Host "       Leave it empty and press Enter to skip $DisplayName publishing." -ForegroundColor DarkGray
    Write-Host

    $inputToken = Read-Host "Please enter your $DisplayName personal token"
    if (-not $inputToken) { return $null }

    try {
        [Environment]::SetEnvironmentVariable($EnvName, $inputToken, "User")
        [Environment]::SetEnvironmentVariable($EnvName, $inputToken, "Process")
        Write-Host "[OK] Token saved to user environment variable. You won't need to enter it again." -ForegroundColor Green
    } catch {
        [Environment]::SetEnvironmentVariable($EnvName, $inputToken, "Process")
        Write-Host "[WARN] Could not save token persistently. It will only be valid for this session." -ForegroundColor Yellow
    }
    Write-Host

    return $inputToken
}

# Find the git remote that points at the GitHub repo; create one if absent.
function Resolve-GithubRemoteName {
    foreach ($line in (& git remote -v 2>$null)) {
        if ($line -match "^(\S+)\s+\S*github\.com[/:]$GithubOwner/$GithubRepo(\.git)?\s+\(fetch\)") {
            return $Matches[1]
        }
    }
    & git remote add github $GithubRepoUrl 2>&1 | Out-Null
    Write-Host "  [INFO] Added git remote 'github' -> $GithubRepoUrl" -ForegroundColor Gray
    return "github"
}

# ============================================
# 0. Pre-checks
# ============================================
Write-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  wxd.files.winforms Auto Release Tool" -ForegroundColor Cyan
Write-Host "  (dual publish: Gitee + GitHub)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host

$GiteeToken = Resolve-ApiToken -EnvName "GITEE_TOKEN" -DisplayName "Gitee" -ScopeHint "projects"
if (-not $GiteeToken) {
    Write-Host "[ERROR] Gitee token cannot be empty. Cancelled." -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

# GitHub is optional: without a token the script still publishes to Gitee only.
$GithubToken = Resolve-ApiToken -EnvName "GITHUB_TOKEN" -DisplayName "GitHub" -ScopeHint "repo (classic token)"

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
    # ★ 自动计算下一个版本号（patch 自增）
    $suggestedVersion = ""
    if ($currentTag -match '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
        $parts = $currentTag -split '\.'
        $nextPatch = [int]$parts[3] + 1
        $suggestedVersion = "$($parts[0]).$($parts[1]).$($parts[2]).$nextPatch"
    }

    if ($currentTag) {
        Write-Host "Current version: $currentTag" -ForegroundColor Yellow
    }
    if ($suggestedVersion) {
        Write-Host "Suggested next version: $suggestedVersion (press Enter to accept)" -ForegroundColor Green
    }

    $Version = Read-Host "Enter new version number (e.g. 1.0.2.4)"

    # ★ 如果用户直接回车且存在建议版本号，使用建议版本号
    if (-not $Version -and $suggestedVersion) {
        $Version = $suggestedVersion
        Write-Host "Using suggested version: $Version" -ForegroundColor Green
    }
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
if ($GithubToken) {
    Write-Host "  Targets : Gitee + GitHub" -ForegroundColor White
} else {
    Write-Host "  Targets : Gitee only (no GitHub token)" -ForegroundColor Yellow
}
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
Write-Host "[1/8] Building project (dotnet publish)..." -ForegroundColor Cyan
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
    "-p:AssemblyVersion=$Version",
    "-p:FileVersion=$Version",
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

Write-Host "[1/8] Build successful!" -ForegroundColor Green

# ============================================
# 5. Locate published EXE (single-file mode)
# ============================================
Write-Host
Write-Host "[2/8] Locating published EXE..." -ForegroundColor Cyan

$ExeName = "wxd.files.winforms.exe"
$ExePath = Join-Path $PublishDir $ExeName

if (-not (Test-Path $ExePath)) {
    Write-Host "[ERROR] Published EXE not found: $ExePath" -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

$exeSize = [math]::Round((Get-Item $ExePath).Length / 1MB, 2)
Write-Host "[2/8] EXE found: $ExeName ($exeSize MB)" -ForegroundColor Green

# ============================================
# 6. Generate MD5 (of EXE, not ZIP)
# ============================================
Write-Host
Write-Host "[3/8] Generating MD5 of EXE..." -ForegroundColor Cyan

$md5 = (Get-FileHash $ExePath -Algorithm MD5).Hash.ToLower()

# ★ 校验 MD5 有效性（32位十六进制）
if ([string]::IsNullOrEmpty($md5) -or $md5.Length -ne 32) {
    Write-Host "[ERROR] MD5 computation failed or invalid: '$md5'" -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit 1
}

Write-Host "[3/8] MD5 generated: $md5" -ForegroundColor Green

# ============================================
# 7. Create Git Tag, push master + tag to both remotes
# ============================================
Write-Host
Write-Host "[4/8] Creating Git Tag..." -ForegroundColor Cyan

# Create tag only (no commit of md5.txt)
git tag -a $TagName -m "$ReleaseNotes" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[WARN] Tag creation may have failed (may already exist), continuing..." -ForegroundColor Yellow
}

$GithubRemoteName = $null
if ($GithubToken) {
    $GithubRemoteName = Resolve-GithubRemoteName
}

Write-Host "  Pushing master + tag to Gitee ($GiteeRemoteName)..."
git push $GiteeRemoteName master 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
git push $GiteeRemoteName $TagName 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

if ($GithubRemoteName) {
    Write-Host "  Pushing master + tag to GitHub ($GithubRemoteName)..."
    git push $GithubRemoteName master 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    git push $GithubRemoteName $TagName 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
}

Write-Host "[4/8] Tag pushed" -ForegroundColor Green

# ============================================
# 8. Create Gitee Release via API
# ============================================
Write-Host
Write-Host "[5/8] Creating Gitee Release..." -ForegroundColor Cyan

# ★ 双引号串里反引号是转义符，字面反引号要写成 ``；MD5 必须保持 `xxx` 反引号包裹，
#   已发布的旧客户端只认这种写法，去掉反引号会让它解析不到校验值而拒绝更新
$releaseTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$releaseBody = "## $TagName`n`n**Release Time**: $releaseTime`n`n$ReleaseNotes`n`n---`n`n**MD5**: ``${md5}```n**File**: $ExeName`n**Size**: $exeSize MB"

# ★ 使用 JavaScriptSerializer 替代 ConvertTo-Json，避免中文字符被转义为 \uXXXX 导致乱码
Add-Type -AssemblyName System.Web.Extensions
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$apiBody = $serializer.Serialize(@{
    access_token     = $GiteeToken
    tag_name         = $TagName
    name             = $TagName
    body             = $releaseBody
    target_commitish = "master"
    prerelease       = $false
})

$utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes($apiBody)

try {
    $releaseResponse = Invoke-RestMethod `
        -Uri "https://gitee.com/api/v5/repos/$GiteeOwner/$GiteeRepo/releases" `
        -Method Post `
        -Body $utf8Bytes `
        -ContentType "application/json; charset=utf-8"

    $releaseId = $releaseResponse.id
    # Fallback: if html_url is empty, construct it manually
    $releaseUrl = if ($releaseResponse.html_url) { 
        $releaseResponse.html_url 
    } else { 
        "https://gitee.com/$GiteeOwner/$GiteeRepo/releases/$TagName" 
    }
    Write-Host "[5/8] Release created successfully!" -ForegroundColor Green
    Write-Host "  Release ID : $releaseId" -ForegroundColor Gray
    Write-Host "  URL        : $releaseUrl" -ForegroundColor Cyan
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host "[ERROR] Create Release failed: $errorMsg" -ForegroundColor Red

    if ($errorMsg -match "401" -or $errorMsg -match "403" -or $errorMsg -match "Unauthorized") {
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
Write-Host "[6/8] Uploading EXE as attachment..." -ForegroundColor Cyan

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
        Write-Host "[6/8] EXE uploaded successfully!" -ForegroundColor Green
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
# 10. Publish the same release to GitHub (skipped without token)
# ============================================
$githubReleaseUrl = $null

if (-not $GithubToken) {
    Write-Host
    Write-Host "[7/8] [8/8] GitHub publishing skipped (no GITHUB_TOKEN)." -ForegroundColor Yellow
} else {
    # api.github.com 只接受 TLS 1.2，PS 5.1 默认协议可能更低
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $ghApi = "https://api.github.com/repos/$GithubOwner/$GithubRepo"
    $ghHeaders = @{
        "Authorization" = "token $GithubToken"
        "Accept"        = "application/vnd.github+json"
        "User-Agent"    = "wxd.files.winforms release script"
    }

    Write-Host
    Write-Host "[7/8] Creating GitHub Release..." -ForegroundColor Cyan

    $ghPayload = $serializer.Serialize(@{
        tag_name         = $TagName
        name             = $TagName
        body             = $releaseBody
        target_commitish = "master"
        prerelease       = $false
        draft            = $false
    })
    $ghBytes = [System.Text.Encoding]::UTF8.GetBytes($ghPayload)

    $ghReleaseId = $null
    try {
        # 先查该 Tag 是否已有 Release：存在就改正文复用 id，脚本重跑不会被 422 挡住
        $ghRelease = $null
        try {
            $ghRelease = Invoke-RestMethod -Uri "$ghApi/releases/tags/$TagName" -Method Get -Headers $ghHeaders
        } catch {
            $ghRelease = $null
        }

        if ($ghRelease) {
            $ghReleaseId = $ghRelease.id
            Invoke-RestMethod -Uri "$ghApi/releases/$ghReleaseId" -Method Patch `
                -Body $ghBytes -ContentType "application/json; charset=utf-8" -Headers $ghHeaders | Out-Null
            Write-Host "[7/8] GitHub Release updated (already existed for this tag)." -ForegroundColor Green
        } else {
            $ghRelease = Invoke-RestMethod -Uri "$ghApi/releases" -Method Post `
                -Body $ghBytes -ContentType "application/json; charset=utf-8" -Headers $ghHeaders
            $ghReleaseId = $ghRelease.id
            Write-Host "[7/8] GitHub Release created successfully!" -ForegroundColor Green
        }

        $githubReleaseUrl = if ($ghRelease.html_url) {
            $ghRelease.html_url
        } else {
            "https://github.com/$GithubOwner/$GithubRepo/releases/tag/$TagName"
        }
        Write-Host "  Release ID : $ghReleaseId" -ForegroundColor Gray
        Write-Host "  URL        : $githubReleaseUrl" -ForegroundColor Cyan
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Host "[WARN] GitHub Release creation failed: $errorMsg" -ForegroundColor Yellow
        if ($errorMsg -match "401|403|Bad credentials") {
            Write-Host "  Please check that GITHUB_TOKEN is correct and has 'repo' scope." -ForegroundColor Yellow
        }
    }

    if ($ghReleaseId) {
        Write-Host
        Write-Host "[8/8] Uploading EXE to GitHub Release..." -ForegroundColor Cyan

        $responseFile = Join-Path $env:TEMP "github_upload_response.txt"

        try {
            # 同名附件会 422，先删旧的再传，保证重复执行也能得到最新文件
            $existingAssets = @(Invoke-RestMethod -Uri "$ghApi/releases/$ghReleaseId/assets" -Method Get -Headers $ghHeaders)
            foreach ($asset in $existingAssets) {
                if ($asset.name -eq $ExeName) {
                    Invoke-RestMethod -Uri "$ghApi/releases/assets/$($asset.id)" -Method Delete -Headers $ghHeaders | Out-Null
                }
            }

            $ghUploadUrl = "https://uploads.github.com/repos/$GithubOwner/$GithubRepo/releases/$ghReleaseId/assets?name=$ExeName"

            $httpCode = & curl.exe -s -w "%{http_code}" -X POST $ghUploadUrl `
                -H "Authorization: token $GithubToken" `
                -H "Content-Type: application/octet-stream" `
                --data-binary "@$ExePath" `
                -o $responseFile 2>&1

            $responseBody = ""
            if (Test-Path $responseFile) {
                $responseBody = Get-Content $responseFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                Remove-Item $responseFile -Force -ErrorAction SilentlyContinue
            }

            if ($httpCode -eq "201") {
                Write-Host "[8/8] EXE uploaded to GitHub successfully!" -ForegroundColor Green
            } else {
                Write-Host "[WARN] GitHub EXE upload failed (HTTP $httpCode)" -ForegroundColor Yellow
                Write-Host "  Response: $responseBody" -ForegroundColor Gray
                Write-Host "  You can upload $ExeName manually on the Release page later." -ForegroundColor Yellow
            }
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Host "[WARN] GitHub EXE upload failed: $errorMsg" -ForegroundColor Yellow
            Write-Host "  You can upload $ExeName manually on the Release page later." -ForegroundColor Yellow
        }
    }
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
Write-Host "  File    : $ExeName (${exeSize} MB)" -ForegroundColor White
Write-Host "  Gitee   : $releaseUrl" -ForegroundColor Cyan
if ($githubReleaseUrl) {
    Write-Host "  GitHub  : $githubReleaseUrl" -ForegroundColor Cyan
} else {
    Write-Host "  GitHub  : not published" -ForegroundColor Yellow
}
Write-Host "========================================" -ForegroundColor Green
Write-Host

Read-Host "Press any key to exit..."
