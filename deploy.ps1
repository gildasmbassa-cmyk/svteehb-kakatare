# DeployBot - EduPilot Cameroun
# Usage : .\deploy.ps1 "message de commit"
#         .\deploy.ps1 "message" -SkipVercel
#         .\deploy.ps1 -Help

param(
    [string]$Message = "",
    [switch]$SkipVercel,
    [switch]$Help
)

$PROJECT  = "C:\Users\ABC\svteehb-kakatare"
$SRC      = "$PROJECT\src\App.jsx"
$MIN_SIZE = 900000
$PROD_URL = "https://svteehblykama.vercel.app"

function Step { param($t) Write-Host "`n>> $t" -ForegroundColor Cyan }
function Ok   { param($t) Write-Host "  OK  $t" -ForegroundColor Green }
function Fail { param($t) Write-Host "  ERREUR: $t" -ForegroundColor Red; exit 1 }
function Info { param($t) Write-Host "  INFO: $t" -ForegroundColor Yellow }

if ($Help) {
    Write-Host "DeployBot - EduPilot Cameroun"
    Write-Host "Usage: .\deploy.ps1 `"message`""
    Write-Host "       .\deploy.ps1 `"message`" -SkipVercel"
    exit 0
}

Write-Host "`n=== DeployBot - EduPilot Cameroun ===" -ForegroundColor Magenta

# 1. Trouver App.jsx dans Downloads
Step "Recherche App.jsx dans Downloads..."
$downloads   = "$env:USERPROFILE\Downloads"
$source      = Get-ChildItem $downloads -Filter "App*.jsx" |
               Where-Object { $_.Length -gt $MIN_SIZE } |
               Sort-Object LastWriteTime -Descending |
               Select-Object -First 1

if (-not $source) {
    Get-ChildItem $downloads -Filter "App*.jsx" |
        ForEach-Object { Info "$($_.Name) - $([math]::Round($_.Length/1KB)) KB" }
    Fail "Aucun App*.jsx valide (min 900 KB) dans Downloads. Retelechargez depuis le chat."
}
Ok "$($source.Name) - $([math]::Round($source.Length/1KB)) KB"

# 2. Verifier encodage
Step "Verification encodage UTF-8..."
$bytes    = [System.IO.File]::ReadAllBytes($source.FullName)
$corrupt  = $false
for ($i = 0; $i -lt ($bytes.Length - 1); $i++) {
    if ($bytes[$i] -eq 0xC3 -and $bytes[$i+1] -eq 0x83) { $corrupt = $true; break }
}
if ($corrupt) { Fail "Fichier corrompu (mojibake). Retelechargez depuis le chat." }
Ok "Encodage propre"

# 3. Copier
Step "Copie vers src\App.jsx..."
Set-Location $PROJECT
Copy-Item $source.FullName $SRC -Force
Ok "Copie effectuee"

# 4. Build
Step "Build npm..."
$out = npm run build 2>&1
if ($LASTEXITCODE -ne 0) { Write-Host $out; Fail "Build echoue." }
Ok "Build reussi"

# 5. Git
Step "Git commit..."
if ($Message -eq "") {
    $Message = "deploy: EduPilot $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
}
git add src\App.jsx
git commit -m $Message
git push
if ($LASTEXITCODE -ne 0) { Fail "Git push echoue." }
Ok "Push GitHub OK"

# 6. Vercel
if (-not $SkipVercel) {
    Step "Deploiement Vercel..."
    npx vercel --prod
    if ($LASTEXITCODE -ne 0) { Fail "Deploiement Vercel echoue." }
    Ok "Deploye sur $PROD_URL"
} else {
    Info "Vercel ignore (-SkipVercel)"
}

Write-Host "`n=== Deploiement termine ===" -ForegroundColor Green
Write-Host "    $PROD_URL`n"
