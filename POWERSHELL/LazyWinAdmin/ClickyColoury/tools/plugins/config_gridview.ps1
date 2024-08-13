# api: multitool
# version: 0.0
# title: set GridView
# description: Default table output Out-GridView
# type: inline
# category: config
# hidden: 1
# key: gv|gridview
# config: -
#
#
# not yet enabled
# · depends on implementation in CLI version


#Set-Alias Out Out-GridView
#Set-Alias Out Format-Table


#-- just use config variable for now
if ($global:cfg.gridview -match ".*Table.*") {
    $global:cfg.gridview = "Out-GridView"
}
else {
    $global:cfg.gridview = "Format-Table"
}
