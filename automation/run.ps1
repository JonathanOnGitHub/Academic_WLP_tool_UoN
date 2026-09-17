# Convenience launcher - activates the venv and runs the automator.
# Equivalent of run.sh for Windows PowerShell.
$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$Here\.venv\Scripts\Activate.ps1" | Out-Null
& "$Here\.venv\Scripts\python.exe" "$Here\wlp_automator.py" @args
