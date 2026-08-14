@echo off
rem ASCII only on purpose. Putting Chinese text (or chcp) in a .bat makes cmd
rem misparse the following lines, because it reads batch files by byte offset.
cd /d "%~dp0"

rem Prefer pythonw.exe from the FinanceEnv environment: starts with no console window.
set "PYW="
for %%P in (
    "%USERPROFILE%\anaconda3\envs\FinanceEnv\pythonw.exe"
    "%USERPROFILE%\miniconda3\envs\FinanceEnv\pythonw.exe"
    "%LOCALAPPDATA%\anaconda3\envs\FinanceEnv\pythonw.exe"
    "%LOCALAPPDATA%\miniconda3\envs\FinanceEnv\pythonw.exe"
    "%ProgramData%\anaconda3\envs\FinanceEnv\pythonw.exe"
    "C:\anaconda3\envs\FinanceEnv\pythonw.exe"
) do if not defined PYW if exist %%P set "PYW=%%~P"

if defined PYW (
    start "" "%PYW%" "%~dp0app.py"
    exit /b 0
)

rem Fallback: let conda locate the environment. A console window stays open here.
where conda >nul 2>&1
if errorlevel 1 goto noconda
call conda run -n FinanceEnv --no-capture-output python app.py
if errorlevel 1 goto failed
exit /b 0

:noconda
echo.
echo conda was not found on PATH.
echo Open "Anaconda Prompt" and run this file from there.
echo.
pause
exit /b 1

:failed
echo.
echo The app failed to start. Check that the environment exists:
echo     conda env list
echo.
echo See data\error.log for details.
echo.
pause
exit /b 1
