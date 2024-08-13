# api: multitool
# version: 1.2
# title: Current user on PC
# description: Get currently logged on user (via PSLoggedOn.exe)
# type: psexec
# category: cmd
# depends: funcs_base
# img: users.png
# hidden: 0
# key: 6|cu|current|currentuser|loggedon|psloggedon
# config: -
#
# Formerly used PsLoggedOn.exe
#  - now using WMI Win32_user as defined in funcs_base

Param($machine = (Read-Host "Computer"))

Get-CurrentUser $machine
