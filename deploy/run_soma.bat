@echo off
setlocal

cd /d "%~dp0.."

set ENV_FILE=%cd%\deploy\config.env

.\soma_run.exe

endlocal