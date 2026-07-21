Set-Location $PSScriptRoot

Write-Host "`n=== Pushing update ===" -ForegroundColor Cyan

@("index.lock","HEAD.lock") | ForEach-Object {
    $f = ".git\$_"
    if (Test-Path $f) { Remove-Item $f -Force; Write-Host "Removed $f" -ForegroundColor Yellow }
}

git add -A
git commit -m "Add Vault tab to Owner role for Excel export"
git push

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== Done! Vercel is redeploying. ===" -ForegroundColor Green
} else {
    Write-Host "`nPush failed." -ForegroundColor Red
}

pause
