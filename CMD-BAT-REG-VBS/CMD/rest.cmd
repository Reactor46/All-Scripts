@echo off

@powershell -NonInteractive -NoProfile -ExecutionPolicy Unrestricted -Command "& {.\SQL\rest.ps1 %*; exit $LastExitCode }"
exit /B %errorlevel%