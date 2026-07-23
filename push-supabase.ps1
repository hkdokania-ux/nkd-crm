Set-Location $PSScriptRoot

Write-Host "`n=== Pushing update ===" -ForegroundColor Cyan

@("index.lock","HEAD.lock") | ForEach-Object {
    $f = ".git\$_"
    if (Test-Path $f) { Remove-Item $f -Force; Write-Host "Removed $f" -ForegroundColor Yellow }
}

Write-Host "Installing packages..." -ForegroundColor Cyan
npm install

git add -A
git commit -m "8 features: enquiry date cap, model search, DOB warning, auto-caps, chassis stock+OTHERS, exchange value fix, hide finance bank, salesman pending tab + canApprove"
git push

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== Done! Vercel is redeploying. ===" -ForegroundColor Green
} else {
    Write-Host "`nPush failed." -ForegroundColor Red
}

pause
