@echo off
cd /d C:\transferin

:: Check if Git is installed
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Git is not installed or not in PATH.
    exit /b
)

:: Check if this is a Git repository
git rev-parse --is-inside-work-tree >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: C:\transferin is not a Git repository.
    exit /b
)

:: Check for changes before committing
git status --porcelain | findstr /r "^.." >nul
IF %ERRORLEVEL% NEQ 0 (
    exit /b
)

:: Add, commit, and push changes
git add .
git commit -m "Auto commit - %DATE% %TIME%"
git push origin main
exit /b
