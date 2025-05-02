# api: multitool
# type: init-gui
# version: 0.6
# title: UserTools
# description: register .\UserTools\
# hidden: 1
# category: edit
# config: { name: cfg.usertools, value: .\UserTools, type: str, description: shortcut source directory }
#
# Each entry copies the respective file to the user desktop
# (Requires `hostname` and `username` to be set.)

if ((!$CLI) -and (!$e) -and ($GUI.w) -and ($m = $GUI.w.findName("Menu_USERTOOLS"))) {
    $m = $m.Items
    
    if (!($cfg.usertools)) {
       $cfg.usertools = ".\UserTools"
    }

    #-- extract description file
    $desc = Import-CSV ".\data\usertools.description.csv"

    #-- add menu entries
    ForEach ($file in (dir "$($cfg.usertools)\*.*")) {
        $fn = $file.Name

        #-- MenuItem + Image
        $meta = $desc | ? { $_.file -eq $fn }
        if (!$meta) { $meta = @{icon="icon.copy.png"; desc=""} }
        $ICON = W Image @{Source=(Get-IconPath $meta.icon); Height=18; Width=18}
        $m.Add((W MenuItem @{Header="_$fn"; Tooltip=$meta.desc; Icon=$ICON; Add_click={ copy-UserTool "$fn" }.getnewclosure()}))
    }

    #-- handler
    function Copy-UserTool($fn) {
        Out-Gui -b Blue -f Cyan "User Tools install                          "
        $machine = $GUI.machine.Text
        $username = $GUI.username.Text

        #-- verify host+user
        if (!$machine) {
            Out-Gui -f Red "No hostname given"
        }
        else {
            if ((!$username) -or (!(Test-Path "\\$machine\c$\Users\$username\Desktop\"))) {
                $username = "Public"
            }
            $dest = "\\$machine\c$\Users\$username\Desktop\"

            #-- copy
            Out-Gui "❏ '$fn' to $dest"
            Copy "$($cfg.usertools)\$fn" "$dest"
        }
    }
}
