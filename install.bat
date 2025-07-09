@echo off
chcp 65001 >nul
echo Installing Outlook MCP Server...
echo.

echo 1. Installing Node.js dependencies...
call npm install
if %errorlevel% neq 0 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo 2. Compiling TypeScript code...
call npm run build
if %errorlevel% neq 0 (
    echo Error: Compilation failed
    pause
    exit /b 1
)

echo.
echo 3. Checking if Outlook is running...
tasklist /FI "IMAGENAME eq OUTLOOK.EXE" 2>NUL | find /I /N "OUTLOOK.EXE">NUL
if "%ERRORLEVEL%"=="0" (
    echo OK: Outlook is running
) else (
    echo Warning: Outlook is not running, please start Outlook first
)

echo.
echo 4. Testing PowerShell permissions...
powershell -Command "Get-ExecutionPolicy" >nul 2>&1
if %errorlevel% neq 0 (
    echo Warning: PowerShell execution policy may be restricted
    echo If you encounter permission issues, run as administrator:
    echo Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
)

echo.
echo Installation completed!
echo.
echo Next steps:
echo 1. Make sure Outlook is running
echo 2. Add the following configuration to Claude Desktop config file:
echo.
echo {
echo   "mcpServers": {
echo     "outlook": {
echo       "command": "node",
echo       "args": ["path/to/outlook/dist/index.js"],
echo       "env": {}
echo     }
echo   }
echo }
echo.
echo 3. Restart Claude Desktop
echo 4. Start using Outlook MCP Server!
echo.
pause
