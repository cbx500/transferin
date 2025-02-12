@echo off
cd /d C:\transferin

:: Log file for debugging
set LOGFILE=C:\transferin\git_log.txt

:: Check if Git is installed
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git is not installed or not in PATH. >> %LOGFILE%
    echo ERROR: Git is not installed or not in PATH.
    exit /b
)

:: Check if this is a Git repository
git rev-parse --is-inside-work-tree >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: C:\transferin is not a Git repository. >> %LOGFILE%
    echo ERROR: C:\transferin is not a Git repository.
    exit /b
)

:: Check for changes before committing
git status --porcelain > git_changes.txt
findstr /r "^.." git_changes.txt >nul
IF %ERRORLEVEL% NEQ 0 (
    echo No changes to commit. >> %LOGFILE%
    echo No changes to commit.
    del git_changes.txt
    exit /b
)

:: Add, commit, and push changes
git add .
git commit -m "Auto commit - %DATE% %TIME%" >> %LOGFILE% 2>&1
git push origin main >> %LOGFILE% 2>&1

echo ✅ Changes successfully pushed to GitHub! >> %LOGFILE%
echo ✅ Changes successfully pushed to GitHub!
del git_changes.txt
exit /b
