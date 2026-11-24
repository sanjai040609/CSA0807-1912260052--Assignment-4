@echo off
REM =================================================================
REM Progressive Git Push Utility for Assignment 4 (Windows Batch)
REM Student: V. Karthik Harish (Reg No: 192511328)
REM Repository: CSA08-192511328-Assignment4
REM =================================================================

echo =================================================================
echo CSA08 ASSIGNMENT 4 - GIT REPOSITORY SYNCHRONIZATION
echo Student: V. Karthik Harish (192511328)
echo =================================================================

git status
echo.
echo Commits ready for submission:
git log --oneline
echo.

echo Pushing all progressive task commits to main branch...
git push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo [SUCCESS] Progressive commits synchronized with remote repository.
) else (
    echo.
    echo [NOTE] Push completed or requires credentials. Check git output above.
)
pause
