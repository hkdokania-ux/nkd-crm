@echo off
cd /d %~dp0
echo.
echo === Deploying NKD CRM ===
echo.

del .git\index.lock 2>nul
del .git\HEAD.lock 2>nul

git add -A
git commit -m "Update"
git push

echo.
echo === Done! Vercel is rebuilding. Wait ~1 min then refresh the app. ===
echo.
pause
