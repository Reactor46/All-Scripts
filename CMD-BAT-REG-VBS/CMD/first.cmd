@echo off

@powershell -NonInteractive -NoProfile -ExecutionPolicy Unrestricted -Command "& {.\SQL\first.ps1 %*; exit $LastExitCode }"
exit /B %errorlevel%