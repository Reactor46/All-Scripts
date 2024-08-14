@echo off

rem Update-DeviceConfig.cmd
rem
rem Start Update-DeviceConfig.ps1, bypassing execution policy.

powershell.exe -ExecutionPolicy Bypass %~dp0Update-DeviceConfig.ps1 -PoliciesPath D:\Store\Policies