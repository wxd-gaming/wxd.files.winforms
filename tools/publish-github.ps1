$ErrorActionPreference='Stop'
$nl=[char]10
$input = "protocol=https"+$nl+"host=github.com"+$nl+$nl
$out = $input | git credential fill 2>$null
$pass = @($out | Select-String '^password=' | ForEach-Object { $_.Line.Substring(9) })[0]
if (-not $pass) { throw 'GitHub credential not found' }

$api='https://api.github.com/repos/wxd-gaming/wxd.files.winforms'
$headers=@{ Authorization=("token " + $pass); Accept='application/vnd.github+json'; 'User-Agent'='wxd-release' }
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$notes = Get-Content -Raw -Encoding UTF8 'tools\release-notes-1.0.2.67.md'
$body = "## v1.0.2.67" + $nl + $nl + $notes + $nl + $nl + "---" + $nl + $nl + "**MD5**: 8843ae9b2ec511f1ee71283cef960b42" + $nl + "**File**: wxd.files.winforms.exe" + $nl + "**Size**: 3.13 MB"

$release=$null
try { $release = Invoke-RestMethod -Uri "$api/releases/tags/v1.0.2.67" -Headers $headers -TimeoutSec 30 } catch { }

$payload = @{ tag_name='v1.0.2.67'; target_commitish='master'; name='v1.0.2.67'; body=$body; draft=$false; prerelease=$false } | ConvertTo-Json -Compress
$bytes=[Text.Encoding]::UTF8.GetBytes($payload)

if ($release) {
  $release = Invoke-RestMethod -Uri "$api/releases/$($release.id)" -Method Patch -Headers $headers -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 30
  "updated_release=$($release.id)"
} else {
  $release = Invoke-RestMethod -Uri "$api/releases" -Method Post -Headers $headers -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 30
  "created_release=$($release.id)"
}

foreach ($a in @($release.assets)) {
  if ($a.name -eq 'wxd.files.winforms.exe') {
    Invoke-RestMethod -Uri "$api/releases/assets/$($a.id)" -Method Delete -Headers $headers -TimeoutSec 30 | Out-Null
    "deleted_old_asset=$($a.id)"
  }
}

$exe = (Resolve-Path 'publish\wxd.files.winforms.exe').Path
Add-Type -AssemblyName System.Net.Http
$client = New-Object System.Net.Http.HttpClient
$client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('token', $pass)
$client.DefaultRequestHeaders.UserAgent.ParseAdd('wxd-release')
$bytes = [IO.File]::ReadAllBytes($exe)
$content = [System.Net.Http.ByteArrayContent]::new($bytes)
$content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/octet-stream')
$uploadUrl = "https://uploads.github.com/repos/wxd-gaming/wxd.files.winforms/releases/$($release.id)/assets?name=wxd.files.winforms.exe"
$resp = $client.PostAsync($uploadUrl, $content).GetAwaiter().GetResult()
"upload_status=" + [int]$resp.StatusCode
$text = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
if ($text.Length -gt 600) { $text = $text.Substring(0,600) }
"upload_body=$text"

$fresh = Invoke-RestMethod -Uri "$api/releases/tags/v1.0.2.67" -Headers $headers -TimeoutSec 30
"html_url=$($fresh.html_url)"
"assets=" + (($fresh.assets | ForEach-Object { $_.name + ':' + $_.size }) -join ',')