@echo off
REM One-time setup for Council: installs the three official CLIs.
REM Needs Node.js 20+ (for Claude Code and Copilot CLI): https://nodejs.org

echo.
echo === Installing Claude Code (claude) ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://claude.ai/install.ps1 | iex"

echo.
echo === Installing GitHub Copilot CLI (copilot) ===
call npm install -g @github/copilot

echo.
echo === Installing Antigravity CLI (agy) - Google's replacement for Gemini CLI ===
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://antigravity.google/cli/install.ps1 | iex"

echo.
echo ============================================================
echo  Done. Now CLOSE this window, open a NEW terminal, and sign in once to each:
echo.
echo    claude     ^(choose your Claude Pro account, then type /exit^)
echo    agy        ^(sign in with the Google account that has Google AI Pro^)
echo    copilot    ^(type /login, sign in with GitHub, then /exit^)
echo.
echo  Then double-click run_council.bat
echo ============================================================
pause
