$date = Get-Date -Format "yyyy-MM-dd_HHmm"
if (!(Test-Path "backups")) { New-Item -ItemType Directory -Path "backups" | Out-Null }
Copy-Item src\App.jsx "backups\App_$date.jsx"
Write-Host "Sauvegarde créée : backups\App_$date.jsx"
vercel --prod
