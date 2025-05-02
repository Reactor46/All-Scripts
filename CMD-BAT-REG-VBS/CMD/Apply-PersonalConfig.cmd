@echo off

rem Apply-PersonalConfig.cmd
rem
rem Start Apply-PersonalConfig.ps1, bypassing execution policy.

powershell.exe -ExecutionPolicy Bypass %~dp0Apply-PersonalConfig.ps1 -StorePath D:\Store