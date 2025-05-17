@echo off
SETLOCAL EnableDelayedExpansion
chcp 65001 > nul

:checkPrivileges
    NET FILE 1>NUL 2>NUL
    IF '%errorlevel%' == '0' ( GOTO :gotPrivileges ) ELSE ( GOTO :requestPrivileges )

:requestPrivileges
    echo Requesting Administrator to run...
    echo If the User Account Control dialog appears, please select "Yes" to continue.
    set "ARG_PATH_FOR_ELEVATION=%~dp0"
    powershell -Command "Start-Process -FilePath '%~f0' -ArgumentList @('%ARG_PATH_FOR_ELEVATION%') -Verb RunAs"
    exit /b

:gotPrivileges
    echo Administrator: Pass!
    
    SET "SCRIPT_DIR_FROM_ARG=%~1"
    SET "CURRENT_BATCH_DIR=%~dp0"
    
    SET "EFFECTIVE_SCRIPT_DIR=%SCRIPT_DIR_FROM_ARG%"
    IF "%EFFECTIVE_SCRIPT_DIR%"=="" SET "EFFECTIVE_SCRIPT_DIR=%CURRENT_BATCH_DIR%"
    
    SET "PS_SCRIPT_NAME=Install-Software.ps1"
    SET "PS_SCRIPT_FULL_PATH=%EFFECTIVE_SCRIPT_DIR%%PS_SCRIPT_NAME%"

    IF NOT EXIST "%PS_SCRIPT_FULL_PATH%" (
        echo ERROR: PowerShell script '%PS_SCRIPT_NAME%' not found at '%PS_SCRIPT_FULL_PATH%'
        echo Please ensure '%PS_SCRIPT_NAME%' is in the same directory as this batch file.
        goto ScriptEnd
    )

    echo File UI: Pass!
    powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT_FULL_PATH%" -ScriptBaseDir "%EFFECTIVE_SCRIPT_DIR%"
    echo PowerShell script execution finished. Exit code: %errorlevel%

:ScriptEnd
ENDLOCAL
echo.
echo Press any key to exit...
pause >nul
exit /b