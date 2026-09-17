@echo off
rem Convenience launcher - activates the venv and runs the automator.
rem Equivalent of run.sh for Windows CMD.
setlocal EnableExtensions EnableDelayedExpansion
set "HERE=%~dp0"
call "%HERE%.venv\Scripts\activate.bat" >nul
"%HERE%.venv\Scripts\python.exe" "%HERE%wlp_automator.py" %*
endlocal
