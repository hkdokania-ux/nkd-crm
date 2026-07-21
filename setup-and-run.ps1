# NKD Bajaj CRM — Setup & Dev Server
# Double-click this file (or right-click → Run with PowerShell) to install deps and start the app.

Set-Location $PSScriptRoot

Write-Host "`n=== NKD Bajaj CRM Setup ===" -ForegroundColor Cyan

# 1. Install dependencies
Write-Host "`n[1/2] Installing npm dependencies..." -ForegroundColor Yellow
npm install
if ($LASTEXITCODE -ne 0) { Write-Host "npm install failed." -ForegroundColor Red; pause; exit 1 }

# 2. Start dev server
Write-Host "`n[2/2] Starting dev server at http://localhost:5173 ..." -ForegroundColor Yellow
Write-Host "      Press Ctrl+C to stop." -ForegroundColor Gray
npm run dev
