@echo off
echo =========================================
echo Running SaveTrigger...
echo =========================================
dotnet run -- %*

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with error code %ERRORLEVEL%.
)
echo.
pause
