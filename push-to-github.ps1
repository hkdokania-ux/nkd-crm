# NKD Bajaj CRM — Push to GitHub
# Run AFTER creating the repo on github.com (see instructions below).
#
# BEFORE running this script:
#   1. Go to https://github.com/new
#   2. Repository name: nkd-crm
#   3. Visibility: Private  (recommended)
#   4. Leave README / .gitignore / license all UNCHECKED
#   5. Click "Create repository"
#   6. Then run this script.

Set-Location $PSScriptRoot

Write-Host "`n=== Push nkd-crm to GitHub ===" -ForegroundColor Cyan

# Add remote (safe to run even if already added)
git remote remove origin 2>$null
git remote add origin https://github.com/hkdokania/nkd-crm.git

Write-Host "`nPushing to github.com/hkdokania/nkd-crm ..." -ForegroundColor Yellow
Write-Host "(A browser login window may appear — sign in to your GitHub account)`n" -ForegroundColor Gray

git push -u origin main

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== Done! ===" -ForegroundColor Green
    Write-Host "Repo live at: https://github.com/hkdokania/nkd-crm" -ForegroundColor Cyan
} else {
    Write-Host "`nPush failed. Make sure you created the repo on GitHub first (see instructions at top of this file)." -ForegroundColor Red
}

pause
