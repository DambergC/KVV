@echo off
setlocal

rem Parameters for flexibility
set "LocalPath=%APPDATA%\Microsoft\Signatures"
set "BackupPath=%APPDATA%\Backup\Signatures"
set "LogFilePath=%BackupPath%\BackupRestore.log"

rem Function to handle log rotation
:RotateLog
if exist "%LogFilePath%" (
    for /f "tokens=1,2 delims=." %%a in ('dir /b /o-n "%LogFilePath%.*" 2^>nul') do (
        set "oldLog=%LogFilePath%.%%a"
        set "newLog=%LogFilePath%.%%b"
        if exist "!oldLog!" ren "!oldLog!" "!newLog!"
    )
    ren "%LogFilePath%" "%LogFilePath%.1"
)
echo. > "%LogFilePath%"
goto :eof

rem Logging utility for better feedback
:LogMessage
set "Timestamp=%date% %time%"
echo [%Timestamp%] [%2] %1 >> "%LogFilePath%"
echo [%Timestamp%] [%2] %1
goto :eof

rem Example usage of LogMessage in the script
call :LogMessage "Script started." "Info"

rem Function to perform copy
:PerformCopy
if not exist "%2" mkdir "%2"
xcopy "%1" "%2" /E /I /Y >nul
call :LogMessage "%3: %1 -> %2" "Info"
goto :eof

rem Check if local signatures exist
if exist "%LocalPath%" (
    rem Check if backup signatures exist
    if exist "%BackupPath%" (
        rem Compare files and decide whether to backup or restore
        for /r "%LocalPath%" %%f in (*) do (
            set "relativePath=%%~pnxf"
            set "relativePath=!relativePath:%LocalPath%=!"
            set "backupFile=%BackupPath%!relativePath!"
            if exist "!backupFile!" (
                rem Compare timestamps
                for %%a in ("%%f") do set "localTimestamp=%%~ta"
                for %%a in ("!backupFile!") do set "backupTimestamp=%%~ta"
                if "!localTimestamp!" gtr "!backupTimestamp!" (
                    call :PerformCopy "%%f" "!backupFile!" "Update backup for"
                ) else if "!localTimestamp!" lss "!backupTimestamp!" (
                    call :PerformCopy "!backupFile!" "%%f" "Restore"
                )
            ) else (
                call :PerformCopy "%%f" "!backupFile!" "Backup"
            )
        )
    ) else (
        rem Backup local signatures if no backup exists
        call :PerformCopy "%LocalPath%\*" "%BackupPath%" "Backup all local signatures"
    )
) else (
    rem Restore from backup if no local signatures exist
    if exist "%BackupPath%" (
        call :PerformCopy "%BackupPath%\*" "%LocalPath%" "Restore all signatures from backup"
    ) else (
        call :LogMessage "No local signatures or backup found. Nothing to do." "Error"
    )
)

pause
