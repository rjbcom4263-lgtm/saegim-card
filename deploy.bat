@echo off
del /f /q ".git\index.lock" 2>nul
del /f /q ".git\HEAD.lock" 2>nul
del /f /q ".git\REBASE_HEAD.lock" 2>nul
del /f /q ".git\MERGE_HEAD.lock" 2>nul
del /f /q ".git\packed-refs.lock" 2>nul
git add -A
git commit -m "%~1"
git push
firebase deploy --only hosting,firestore:rules,storage
echo.
echo Done!
