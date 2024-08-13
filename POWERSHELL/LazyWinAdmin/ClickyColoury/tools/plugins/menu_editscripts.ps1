# api: multitool
# type: init-gui
# version: 0.5
# title: --- edit ---
# description: edit multitool tools/ and main script, modules
# hidden: 1
# category: edit
# depends: wpf
# config: {}
#
# Adds to Config > Edit scripts menu

if ((!$CLI) -and (!$e) -and ($GUI.w)) {

    $wm = $GUI.w.findName("Menu_EDIT").Items
    $submenus = @{}

    #-- prepend main scripts
    $add = @(
       @{fn = ".\modules\starter.ps1"}
       @{fn = ".\modules\wpf.psm1"},
       @{fn = ".\modules\wpf.xaml"}
       @{fn = ".\modules\menu.psm1"}
       @{fn = ".\modules\clipboard.psm1"}
    )

    #-- add edit entries
    ($add+$menu) | ? {$_.fn} | Sort-Object {$_.fn} | % {

        #-- dir and path
        if (($_fn = $_.fn) -match "([.\w]+)[\\//]([^\\//]+)$") {
            $dir = $matches[1] -replace "tools\.",""
            $fn = $matches[2]
        }
        else {
            continue
        }

        #-- find/add dir submenu
        if ($m = $submenus[$dir]) {
        }
        else {
            $m = W MenuItem @{Header=$dir}
            $submenus[$dir] = $m
            $wm.Add($m)
        }
        
        $m.Items.Add((W MenuItem @{Header="_$fn"; Add_click={notepad "$_fn"}.getnewclosure()}))
    }
}
