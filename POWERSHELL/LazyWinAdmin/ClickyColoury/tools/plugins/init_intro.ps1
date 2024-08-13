# api: ps
# title: init screen
# description: writes something on the Output window on startup
# version: 0.0.3
# type: init-gui
# depends: wpf
# category: misc
# hidden: 1
# priority: core

if ($GUI.w -and !$CLI) {

    # combine version of all scripts
    $sigma_ver = ($menu | % { $_.version } | ? { $_ -match "\d" } | % { $_ -replace "-.+$","" -replace "\D","" } | Measure -Sum).Sum -replace "(?<=\d)(?!$)","."

    # output
    Out-Gui -f Yellow -b "#223388" "ClickyColoury frontend to Multi-Tools  Σ ≈ $sigma_ver"
    Out-Gui -f "#88bb22" " ✉ Color clipboard enabled (➤HTML icon)"
}
