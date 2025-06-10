@echo off
title MFA Self-Check Logger

:: Identity Check
set /p usercheck=Are you Jason Santiago? [yes/no]: 
if /i not "%usercheck%"=="yes" (
    echo You are not authorized to run this script.
    pause
    exit /b
)

:: Get timestamp
for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set mydate=%%b-%%c-%%d
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set mytime=%%a_%%b

:: Set custom log folder path
set "logfolder=C:\Users\redacted\Documents\_Bakup2024-2025\MFA_Log"
if not exist "%logfolder%" (
    mkdir "%logfolder%"
)

:: Set log file path with timestamp
set "logfile=%logfolder%\MFA_Log_%mydate%_%mytime%.txt"

:: Begin log
echo Security+ MFA Self-Check >> "%logfile%"
echo Username: %USERNAME% >> "%logfile%"
echo Timestamp: %date% %time% >> "%logfile%"
echo. >> "%logfile%"

:: Prompt for Security+ MFA components
set /p know=Do you have something you know? (e.g., password) [yes/no]: 
echo Something you know: %know% >> "%logfile%"
if /i not "%know%"=="yes" goto denied

set /p have=Do you have something you have? (e.g., phone or token) [yes/no]: 
echo Something you have: %have% >> "%logfile%"
if /i not "%have%"=="yes" goto denied

set /p are=Do you have something you are? (e.g., fingerprint or face ID) [yes/no]: 
echo Something you are: %are% >> "%logfile%"
if /i not "%are%"=="yes" goto denied

:: NIST Awareness Question
set /p nist=What NIST standard does the MFA authentication practice? 
if /i not "%nist%"=="800-53" if /i not "%nist%"=="SP 800-53" (
    echo Incorrect. The correct NIST standard is SP 800-53.
    echo Exiting without launching the portal...
    echo NIST Awareness Question Failed >> "%logfile%"
    goto end
)

echo NIST Standard confirmed: %nist% >> "%logfile%"
echo MFA Self-Check Passed >> "%logfile%"
echo Launching secure Microsoft portal...
start https://mysignins.microsoft.com/security-info
goto end

:denied
echo MFA Self-Check Failed >> "%logfile%"
echo Secure portal will not be launched.

:end
echo Log saved to: %logfile%
pause