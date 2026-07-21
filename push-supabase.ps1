Set-Location $PSScriptRoot

Write-Host "`n=== Pushing Supabase migration ===" -ForegroundColor Cyan

# Remove stale git lock files
@("index.lock","HEAD.lock") | ForEach-Object {
    $f = ".git\$_"
    if (Test-Path $f) { Remove-Item $f -Force; Write-Host "Removed $f" -ForegroundColor Yellow }
}

git add -A
git commit -m "Migrate from Firebase to Supabase for cloud storage"
git push

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== Done! Vercel is redeploying. ===" -ForegroundColor Green
} else {
    Write-Host "`nPush failed." -ForegroundColor Red
}

pause
