# api: multitool
# version: 0.4
# title: Config file
# description: Create and edit config file
# type: inline
# depends: menu
# category: config
# hidden: 1
# key: config|cfg
# img: tools
# config: -
#
#
# Create/load main config file


#-- init
Param($fn="$env:APPDATA\multitool\config.ps1", $options=@(), $overwrite=0, $CRLF="`r`n", $EDITOR="notepad")

# create parent dir
if ($overwrite -or !(Test-Path $fn)) {
    $dir = Split-Path $fn -parent
    if ($dir -and !(Test-Path $dir)) { md -Force "$dir" }
}

#-- read file
if (Test-Path $fn) {
    $src = (Get-Content $fn) | Out-String
}
else {
    $src = "# type: config$CRLF# fn: $fn$CRLF$CRLF"
}

#-- fetch options from all plugins
$options = @($menu | ? { $_.config -and $_.config.count } | % { $_.config })
#-- and main
$options += (Extract-PluginMeta "./modules/starter.ps1").config
$options += (Extract-PluginMeta "./modules/menu.psm1").config
$options += (Extract-PluginMeta "./modules/wpf.psm1").config

#-- assemble defaults
$options | % {
    $v = $_.value
    switch ($_.type) {
        bool { $v = @('$False', '$True')[[Int32]$v] }
        default { $v = "'$v'" }
    }
    if ($src -notmatch "(?mi)^[$]$($_.name)") {
        $src += '$' + $_.name + " = " + $v + "; # $($_.description)$CRLF"
    }
}

# (over)write
$src | Out-File $fn -Encoding UTF8

#-- start notepad
& $EDITOR "$fn"
