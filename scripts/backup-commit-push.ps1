param(
  [string]$Message = "",
  [string]$RemoteUrl = "https://github.com/dewecorp/qurbanku.git",
  [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
  $PSNativeCommandUseErrorActionPreference = $false
}

function Test-GitCommand {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$GitArgs)

  $oldErrorActionPreference = $ErrorActionPreference
  $ErrorActionPreference = "Continue"
  try {
    & git @GitArgs 1>$null 2>$null
    return ($LASTEXITCODE -eq 0)
  } finally {
    $ErrorActionPreference = $oldErrorActionPreference
  }
}

function Remove-BrokenGitDirectory {
  param([string]$PathToGitDir, [string]$RepoRoot)

  if (-not (Test-Path -LiteralPath $PathToGitDir)) {
    return
  }

  $resolvedGitDir = [System.IO.Path]::GetFullPath($PathToGitDir)
  $resolvedRoot = [System.IO.Path]::GetFullPath($RepoRoot)
  if (-not $resolvedGitDir.StartsWith($resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Lokasi .git tidak aman untuk dihapus: $resolvedGitDir"
  }

  Write-Host "Folder .git kosong/rusak terdeteksi. Menghapus .git lama sebelum git init..."
  Remove-Item -LiteralPath $PathToGitDir -Recurse -Force
}

$defaultMessage = if ($Message) { $Message } else { "Update QurbanKu app" }
$inputMessage = Read-Host "Masukkan nama/pesan commit (Enter untuk: $defaultMessage)"
if (($null -ne $inputMessage) -and $inputMessage.Trim()) {
  $Message = $inputMessage.Trim()
} else {
  $Message = $defaultMessage
}

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $root

Write-Host "Folder kerja: $root"

$backupDir = Join-Path $root "backups"
if (-not (Test-Path $backupDir)) {
  New-Item -ItemType Directory -Path $backupDir | Out-Null
}

$zipPath = Join-Path $backupDir "qurbanku-backup-latest.zip"
$backupItems = @(
  "index.html",
  "Code.gs",
  "Code_simple.gs",
  "qurbanku-blogger-theme.xml",
  "backup-commit-push.bat",
  "script push github qurbanku.txt",
  "scripts"
)

$backupPaths = @()
foreach ($item in $backupItems) {
  $itemPath = Join-Path $root $item
  if (Test-Path $itemPath) {
    $backupPaths += $itemPath
  }
}

Write-Host "Membuat backup ZIP..."
Compress-Archive -Path $backupPaths -DestinationPath $zipPath -Force
Write-Host "Backup dibuat: $zipPath"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw "Git tidak ditemukan di PATH. Install Git for Windows dulu."
}

$gitDir = Join-Path $root ".git"
$isRepo = Test-GitCommand rev-parse --is-inside-work-tree

if (-not $isRepo) {
  Remove-BrokenGitDirectory -PathToGitDir $gitDir -RepoRoot $root
  Write-Host "Inisialisasi repository Git lokal..."
  git init
}

$remoteExists = Test-GitCommand remote get-url origin

if ($remoteExists) {
  git remote set-url origin $RemoteUrl
} else {
  git remote add origin $RemoteUrl
}
Write-Host "Remote origin: $RemoteUrl"

$gitName = git config user.name
if (-not $gitName) {
  git config user.name "dewecorp"
}

$gitEmail = git config user.email
if (-not $gitEmail) {
  git config user.email "dewecorp@users.noreply.github.com"
}

git add index.html Code.gs Code_simple.gs qurbanku-blogger-theme.xml backup-commit-push.bat "script push github qurbanku.txt" scripts backups/qurbanku-backup-latest.zip

$hasChanges = $false
if (-not (Test-GitCommand diff --cached --quiet)) {
  $hasChanges = $true
}

if ($hasChanges) {
  git commit -m $Message
} else {
  Write-Host "Tidak ada perubahan untuk commit."
}

git branch -M $Branch

Write-Host "Status Git sebelum push:"
git status --short --branch

if (Test-GitCommand ls-remote --exit-code --heads origin $Branch) {
  git pull --rebase --autostash origin $Branch
}

Write-Host "Mulai push ke GitHub. Jika diminta login, ikuti popup Git Credential Manager/GitHub."
git push -u origin $Branch

Write-Host "Selesai: backup, commit, dan push ke $RemoteUrl branch $Branch"
