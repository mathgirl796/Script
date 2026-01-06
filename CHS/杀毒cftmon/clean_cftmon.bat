@echo off
setlocal enabledelayedexpansion

:: Set console to UTF-8 encoding
chcp 936 >nul 2>&1

echo ========================================
echo CFTMON Virus Removal Tool
echo ========================================
echo.

:: Step 1: Kill cftmon.exe process
echo [STEP 1] Terminating cftmon.exe process...
taskkill /F /IM cftmon.exe >nul 2>&1
if !ERRORLEVEL! equ 0 (
    echo [SUCCESS] cftmon.exe process terminated successfully
) else (
    echo [INFO] cftmon.exe process is not running
)
echo.


:: Step 2: Remove system and hidden attributes from cftmon.exe
echo [STEP 2] Removing hidden and system attributes from C:\Windows\cftmon.exe...
if exist "C:\Windows\cftmon.exe" (
    attrib "C:\Windows\cftmon.exe" -S -H
    if !ERRORLEVEL! equ 0 (
        echo [SUCCESS] Attributes removed successfully
    ) else (
        echo [ERROR] Failed to remove attributes
    )
) else (
    echo [INFO] C:\Windows\cftmon.exe does not exist
)
echo.

:: Step 3: Delete cftmon.exe virus file
echo [STEP 3] Deleting virus file C:\Windows\cftmon.exe...
if exist "C:\Windows\cftmon.exe" (
    del /F /S /Q "C:\Windows\cftmon.exe"
    if !ERRORLEVEL! equ 0 (
        echo [SUCCESS] Virus file deleted successfully
    ) else (
        echo [ERROR] Failed to delete virus file
    )
) else (
    echo [INFO] C:\Windows\cftmon.exe does not exist
)
echo.

:: Step 4: Scan and display all exe files from specified drive
echo [STEP 4] Scanning exe files...
set /p drive="Enter drive letter: "
if defined drive (
    set "drive=!drive:~0,1!"
    echo You entered drive: !drive!
    echo Checking if drive exists...
    if exist "!drive!:\\" (
        echo [DEBUG] Drive !drive!: exists
        echo Searching in !drive!: drive...
        echo ========== EXE FILES FOUND ==========
        set /a count=0
        set /a deleted=0
        for /f "delims=" %%f in ('dir "!drive!:\*.exe" /s /b 2^>nul') do (
            set /a count+=1
            for %%A in ("%%f") do (
                echo [!count!] File: %%f
                echo      Size: %%~zA Bytes
                if "%%~zA"=="90112" (
                    echo      [DELETING] 90112 Bytes file detected - DELETING...
                    del /F /Q "%%f"
                    set /a deleted+=1
                    echo      [DELETED] File has been deleted
                ) else (
                    echo      [KEPT] File size is not 90112 Bytes
                )
            )
            echo.
        )
        echo ====================================
        echo [DEBUG] Total exe files found: !count!
        echo [DEBUG] Total files deleted: !deleted!
        echo [SUCCESS] Scan completed
    ) else (
        echo [ERROR] Drive !drive!: does not exist
    )
) else (
    echo [ERROR] No drive specified
)
echo.

:: Step 5: Removing hidden and system attributes from all files in !drive!: drive...
echo [STEP 5] Removing hidden and system attributes from all files in !drive!: drive...
attrib "!drive!:\*" -S -H /S /D
echo [SUCCESS] Attributes removed from all files

PAUSE