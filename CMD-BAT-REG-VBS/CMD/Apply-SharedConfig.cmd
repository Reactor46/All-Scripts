@echo off

rem Apply-SharedConfig.cmd
rem
rem Start Apply-SharedConfig.ps1, bypassing execution policy.

powershell.exe -ExecutionPolicy Bypass %~dp0Apply-SharedConfig.ps1 -StorePath D:\Store