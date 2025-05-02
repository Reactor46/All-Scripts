@ECHO OFF
SET ThisScriptsDirectory="\\Contosocorp\share\Shared\IT\DaRT\"
SET PowerShellScriptPath=%ThisScriptsDirectory%system_report.ps1
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%PowerShellScriptPath%""' -Verb RunAs}";
