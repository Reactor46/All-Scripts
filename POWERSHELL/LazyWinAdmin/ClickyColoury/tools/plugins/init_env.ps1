# api: ps
# title: initialize $ENV
# description: predefines some Powershell and system environment variables
# version: 0.5
# type: init
# category: misc
# hidden: 1
# priority: core

#-- powershell
#$Debug = $true
#$ErrorActionPreference = "SilentlyContinue"
#$DebugPreference = "SilentlyContinue"
#$ProgressPreference = "Continue"
#$VerbosePreference = "SilentlyContinue"
$WarningPreference = "Continue"
#$WhatIfPreference = $False
#$ConfirmPreference = "High"
$OFS=" "

#-- environment
if (!$ENV:EDITOR) {
    $ENV:EDITOR = "notepad"
}
if (!$ENV:XDG_CONFIG_HOME) {
    $ENV:XDG_CONFIG_HOME = $ENV:APPDATA
}

#-- create config dir
if (!(Test-Path ($cfg.user_plugins_d))) {
    md ($cfg.user_plugins_d)
}