@echo off
REM Makes the "claude" command work: installs Claude Code if needed,
REM adds it to your PATH, then starts it so you can sign in.
set "CB=%USERPROFILE%\.local\bin"

if not exist "%CB%\claude.exe" (
  echo Claude Code not found - installing it now...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://claude.ai/install.ps1 | iex"
)

if not exist "%CB%\claude.exe" (
  echo.
  echo Install did not finish. Take a screenshot of this window and send it to Claude.
  pause
  exit /b 1
)

powershell -NoProfile -Command "$d = Join-Path $env:USERPROFILE '.local\bin'; $p = [string][Environment]::GetEnvironmentVariable('Path','User'); if (($p -split ';') -notcontains $d) { [Environment]::SetEnvironmentVariable('Path', ($p.TrimEnd(';') + ';' + $d), 'User'); 'Added Claude to your PATH.' } else { 'Claude is already on your PATH.' }"

echo.
echo Starting Claude. Sign in with your Claude Pro account when your browser opens.
echo When you're done, type /exit
echo.
"%CB%\claude.exe"
pause
