@echo off

rem Gather-DeviceConfig.cmd
rem
rem Start Gather-DeviceConfig.ps1, bypassing execution policy.

powershell.exe -ExecutionPolicy Bypass %~dp0Gather-DeviceConfig.ps1 -StorePath D:\Store