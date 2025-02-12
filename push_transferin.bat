@echo off
cd /d C:\transferin

:: Check if Git is installed
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git is not installed or not in PATH.
    echo Install Git from https://git-scm.com/downloads
    pause
    exit /b
)

:: Check if this is a Git repository
git rev-parse --is-inside-work-tree >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: C:\transferin is not a Git repository.
    echo Run 'git init' and set up the remote repository first.
    pause
    exit /b
)

:: Check for changes before committing
git status --porcelain | findstr /r "^.." >nul
IF %ERRORLEVEL% NEQ 0 (
    echo No changes to commit.
    pause
    exit /b
)

:: Add, commit, and push changes
git add .
git commit -m "Manual commit - %DATE% %TIME%"
git push origin main

echo.
echo ✅ Changes successfully pushed to GitHub!
pause
exit /b
