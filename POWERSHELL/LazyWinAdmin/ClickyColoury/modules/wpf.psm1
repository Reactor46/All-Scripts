# api: ps
# title: WPF + WinForms
# description: WinForm and WPF shortcuts, GUI and thread handling
# depends: menu, clipboard, sys:presentationframework, sys:system.windows.forms, sys:system.drawing
# doc: http://fossil.include-once.org/clickycoloury/wiki/W
# version: 1.0.7
# license: MITL
# type: functions
# category: ui
# config:
#   { name: cfg.threading, type: bool, value: 1, description: Enable threading/runspaces for GUI/WPF version }
#   { name: cfg.autoclear, type: int, value: 300, description: Clear output box after N seconds. }
#   { name: cfg.noheader, type: bool, value: 0, description: Disable script info header in output. }
# status: beta
# priority: default
#
# Handles $menu list in WPF window. Creates widgets and menu entries
# for plugin list / $menu entries.
#   · `type: inline` is the implied default, renders output in TextBlock 
#   · `type: cli` or `window` plugins are run in a separate window
#   · `hidden: 1` tools only show up in menus
#   · `keycode:` is used for shortcuts; the CLI `key:` regex ignored
#   · `type: init-gui` plugins are run once during GUI construction
#   · Whereas `type: init` execute in the main/script RunSpace
#
# The responsive UI is achieved through:
#   · A new runspace for the GUI and a trivial message queue in $GUI.tasks.
#   · Main loop simply holds the window open, then executes $GUI.tasks events.
#   · That event queue simply holds the exact entries from $menu.
#   · Pipes them through `Run-GuiTask` (Should ideally be identical to the
#     one in menu.psm1, but needs a few customizations here.)
#
# All widget interactions are confined to the WPF runspace/thread.
#   · WPF interaction through `$GUI` would often hang both processes.
#   · Thus `Out-Gui` not only manages output, but also variable injection.
#
# Scripts/tools should work identically as for the CLI version:
#   · Aliases for `Write-Host` and `Read-Host` should make it transparent.
#   · However simple console output (typically to the stdout stream) will
#     have to be pipe-captured.
#   · Thus there's no guaranteed order for them and direct `Write-Host` calls.
#
# ToDo:
#   · split up into gui.psm1 and wpf.psm1 (actual GUI runspace)


#-- register libs
Add-Type -AN PresentationCore, PresentationFramework, WindowsBase
Add-Type -AN System.Drawing, System.Windows.Forms, Microsoft.VisualBasic
# [System.Windows.Forms.Application]::EnableVisualStyles()
# [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($true)

#-- init vars
$ModuleDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$BaseDir = Split-Path -Parent $ModuleDir
$global:GUI = [hashtable]::Synchronized(@{Host=$Host})
$menu = @()
$global:last_output = 1.5


#-- WPF/WinForms widget wrapper
function W {
    <#
      .SYNOPSIS
         Convenient wrapper to create WPF or WinForms widgets.
      .DESCRIPTION
         Allows to instantiate and modify WPF/XAML or WinForms UI widgets.
         The $prop hash unifies property assignments and widget method calls.
         Notably this implementation is a bit slow, due to this abstraction and
         probing different object trees.          
      .PARAMERER type
         Widget type (e.g. "Button" or "Label") or existing $w widget object.
      .PARAMETER prop
         HashTable listing widget attributes or methods to call:
           @{Color="Red"; Text="Hello World"; Add=$other_widget}
         Property+method names depend on whether WPF or WinForms widgets are used. 
      .PARAMETER type
         Defaults to "Controls" for WPF widgets, but can be set to "Forms" for WF.
      .OUTPUTS
         Returns the instantiated Widget.
      .EXAMPLE
         Create a button:
           $w_btn = W Button @{Content="Text"; Border=2; add_Click=$cb}
         Add child widgets per list:
           $w_grd = W Grid @{Add=$w1, $w2, $w3}
         Or chain creation of multiple widgets:
           $w_all = W StackPanel {Spacing=5; Add=(
              (W Button @{Content="OK"}),
              (W Label @{Content="Really?"})
           )}
         The nesting gets confusing for WPF, but often simplifies WinForm structures.
      .NOTES
         The shortcut method `WF` creates WinForms widgets, whereas `WD` is
         for TextBlock/document inlines.
    #>
    [CmdletBinding()]
    Param($type = "Button", $prop = @{}, $Base="Controls")

    #-- new object
    if ($type.getType().Name -eq "String") {
        $w = New-Object System.Windows.$Base.$type
    }
    else {
        $w = $type      
    }
    #@bug on FOOTERM01 w/ PS 3.0
    if (($PSVersionTable.PSVersion.Major -eq 3) -and ($w -is [System.Windows.Thickness])) { return $w; }

    #-- apply options+methods
    $prop.keys | % {
        $key = $_
        $val = $prop[$_]
        if ($pt = ($w | Get-Member -Force -Name $key)) { $pt = $pt.MemberType }

        #-- properties
        if ($pt -eq "Property") {
            if (($Base -eq "Forms") -and (@("Size" , "ItemSize", "Location") -contains $key)) {
                $val = New-Object System.Drawing.Size($val[0], $val[1])
            }
            $w.$key = $val
        }
        #-- check for methods in widget and common subcontainers
        else {
            ForEach ($obj in @($w, $w.Children, $w.Child, $w.Container, $w.Controls)) {
                if ($obj.psobject -and $obj.psobject.methods.match($key) -and $obj.$key) {
                    ([array]$val) | ? { $obj.$key.Invoke } | % { $obj.$key.Invoke($_) } | Out-Null
                    break
                }
            }
        }
    }
    return $w
}

#-- WinForms version
function WF {
    Param($type, $prop, $add=@{}, $click=$null)
    W -Base Forms -Type $type -Prop ($prop + @{add=$add; add_click=$click})
}

#-- Document "widgets"
function WD {
    Param($type, $prop=@{})
    W -Base Documents -Type $type -Prop $prop
}


#-- WPF main window
function WPF-Window {
    <#
      .SYNOPSIS
         Loads XAML file `wpf.xaml` from same directory as this module.
      .DESCRIPTION
         Also populates $GUI.$_ for a select few global widget names.
      .NOTES
         Workaround img\ cache dir not enabled.
    #>

    #-- ImgDir
    $ImgDir = $BaseDir
    #if (Test-Path "$($env:APPDATA)\multitool\img") { $ImgDir = "$($env:APPDATA)/multitool" }

    #-- load
    $xaml = ((Get-Content "$BaseDir/modules/wpf.xaml" | Out-String) -replace "e:/","$ImgDir/")
    $w = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))

    #-- save aliases int $GUI (sync hash across all threads/runspaces)
    $GUI.w = $w
    $GUI.styles = $w.Resources
    $shortcuts = "Window,Menu,Ribbon,Grid_EXTRAS,Output,machine,username,bulkcsv"
    $shortcuts.split(",") | % { $GUI.$_ = $w.findName($_) } | Out-Null

    #--- return
    return $w
}


#-- create new runspace
function New-GuiThread {
    <#
      .SYNOPSIS
         Creates a new Runspace for the GUI interface.
      .DESCRIPTION
         Copies functions over into new Runspace/Pipeline and attaches shared $GUI
         hashtable for data exchange.
      .EXAMPLE
         New-GuiThread {code} -Funcs "copy1,copy2" -Vars "menu,GUI" -Modules "WPF"
      .NOTES
         Does not work with PS 5.0 yet (presumably scoping or syntax issue with func duplication).
    #>
    Param(
        $code = {},
        $funcs = "WPF-Window,W,WF,WD,Add-GuiMenu,Get-IconPath,Create-GuiParamFields,"+
                 "Out-Gui,Out-Html,Ask-Gui,Add-ButtonHooks,AD-Search,Get-MachineCurrentUser,"+
                 "Set-Clipboard,Set-ClipboardHtml,Get-Clipboard,Get-NotepadCSV,Extract-PluginMeta,preg_match_all",
        $modules = "",
        $vars = "menu",
        $level = "",
        [switch]$invoke=$true
    )

    #-- create, add functions and code
    $shell = [PowerShell]::Create()
    $shell.AddScript((
        ($funcs.split(",")  |  ? { $_ }  |  % { 
            $func = Get-Command $_ -CommandType Filter,Function,Cmdlet
           "$($func.CommandType) $($func.Name) { $($func.Definition) };`r`n"
        }) -join "`r`n"
    )) | Out-Null
#Write-Host -f green "'''$code'''"
    $shell.AddScript($code) | Out-Null

    #-- separate thread
    $shell.Runspace = [runspacefactory]::CreateRunspace()
    $shell.Runspace.ApartmentState = "STA"
    $shell.Runspace.ThreadOptions = "ReuseThread"         
    $shell.Runspace.Open() | out-null

    #-- add vars, modules
    $shell.Runspace.SessionStateProxy.SetVariable("GUI", $GUI) | out-null
    $shell.Runspace.SessionStateProxy.SetVariable("vars", @{}) | out-null
    $shell.Runspace.SessionStateProxy.SetVariable("in_thread", 1) | out-null
    $vars.split(",")  |  ? { $_ } |  % { $shell.Runspace.SessionStateProxy.SetVariable($_, (get-variable $_).value) } | Out-Null
    #$modules.split(",") | ?{$_} | % { $shell.Runspace.SessionStateProxy.ImportPSModule($_) } | Out-Null

    #-- start
    if ($invoke) {
       $handle = $shell.BeginInvoke()
    }
    return $shell
}
#-- no worky
function Attach-StreamEventHandler {
    #-- attaches an event handler to runspace streams (.error and .warning) to capture output
    #$shell.AddScript({
    #    $GUI.stream_error = $Error
    #})
    #Register-ObjectEvent -InputObject $shell.streams.error -EventName DataAdded -Action ({
    #    @{error="err"} | Out-String | Out-Gui -f Red
    #})
    #$shell.streams.error.add_DataAdded({
       # Param($sender, $event)
       #  $Error | Out-Gui -f Red
       #  @($Sender) | out-string | Out-Gui -f Red
       #  [void]($sender.ReadAll() | % { $_.Message | Out-Gui })
       #  Out-Gui "error" -f Red
    #})
}


#-- Look up icon basename path variations
function Get-IconPath {
    <#
      .SYNOPSIS
         Looks up alternative filenames for icons.
      .DESCRIPTION9
         Scans for PNGs (filename from PMD .icon or .img field) in $BaseDir/img/
      .PARAMETER basenames
         Can be a list of icon names "user", "icon.user", "user2.png" etc.
    #>
    [CmdletBinding()]
    Param([Parameter(ValueFromRemainingArguments=$true)]$basenames)
    ForEach ($fn in ($basenames | ? { $_ -ne $null })) {
        ForEach ($png in "icon.$fn.png", "icon.$fn", "$fn.png", "$fn") {
            if (Test-Path ($path = "$BaseDir/img/$png")) {
                return $path
            }
        }
    }
}

#-- Additional input boxes / or combobox if according data/fieldname.txt file exists
function Create-GuiParamFields() {
    <#
      .SYNOPSIS
         Creates input field widgets for extra #param: names
      .DESCRIPTION
         Some plugins may depend on more input than just `machine` and `username`.
         In order to avoid pesky VBS input field popups, plugins may specify additional
         fields with `# param: username,accesslevel,othervar3`

         This function crafts TextBox or ComboBox fields for those. Assigns them a
         unique (per-plugin) widget name "var_$plugin_$paramname". So it can later be read
         out by the "Read-Host" wrapper.
      .NOTES
         Combobox fields are created when there's an according data/combobox.$paramname.txt file.
    #>
    Param(
        $s = "extra,field,names",
        $prefix = "plugin",
        $extra_params = @()
    )
    "$s".split("[;,]") | ? { $_ -and $_ -notmatch "^\s*(machine|host|computer|adname|account|user|bulk)" } | % {
        $key = $_.trim()
        if (Test-Path ($fn="data/combobox.$key.txt")) {
            $field = (W ComboBox @{Width=120; IsEditable="True"})
            Get-Content $fn | % { $field.Items.Add((W ComboBoxItem @{Content=$_})) } | Out-Null
            $field.text = $field.Items[0].content  # use first entry as default
        }
        else {
            $field = (W TextBox @{Width=120})
        }
        $extra_params += (W WrapPanel @{Add=(W Label @{Content="$key"; FontWeight="Bold"}), $field})
        $GUI.w.registerName("var_$($prefix)_$($key)", $field)
    }
    return $extra_params
}

#-- Add tool buttons to main window
function Add-GuiMenu {
    <#
      .SYNOPSIS
         Adds menu entries and button blocks for each plugin from $menu
      .DESCRIPTION
         Iterators over the $menu list
          - skips "hidden:" or "nomenu:" entries, or "type:init*" plugins
          - adds a MenuItem and callback
          - uses the PMD .category to find the right menu or grid/notebook tab
          - looks up plugin or category icons (WPF requires unique instances each?!)
          - calls Create-GuiParamFields for extra input variables
         Also handles shortcut icons section.
      .NOTES
         This is what causes the slow startup! (perhaps due to `W` being too convenient)
    #>
    Param($menu)
    $icon_default = W Image
    ForEach ($e in $menu) {
    
        #-- prepare params
        $e.hidden = ($e.hidden -match "1|yes|hide|hidden|true")
        if ($e.type -eq "init") { continue; }
        $CAT = $e.category.toUpper();
        $GRID = (@($GUI.w.findName("Grid_$CAT"), $GUI.Grid_EXTRAS) -ne $null)[0]

        #-- output category header
        if (($e.category -ne $category) -and (-not $e.hidden)) {
            $category = $e.category
            $GRID.Children.Add((W Label @{Content="  → $category"; Foreground="White"; Background="#443377"; Font="Verdana"; FontSize=17; Width=775}))
        }

        #-- callback (= now just appends to event queue)
        $cb = { $GUI.tasks += $e }.getnewclosure()

        #-- action block/button
        if (-not $e.hidden) {
            $border = W Border @{Style=$GUI.styles.ToolBlock; ToolTip=$e.doc; set_Child=(
                W StackPanel @{Orientation="Horizontal"; Add=
                    (W Button @{Style=$GUI.styles.ToolButton; Width=120; Add_Click=$cb; Content=(
                       W WrapPanel @{Add=
                          (W Image @{Source=(Get-IconPath $e.img $e.icon $e.category); Height=20; Width=22}),
                          (W TextBlock @{Text=$e.title; TextWrapping="Wrap"})
                       }
                    )}),
                    (W StackPanel @{Padding=2; Margin=4; Width=200; Add=
                       @(
                          (W TextBlock @{Text=$e.description; TextWrapping="Wrap"}),
                          (W TextBlock @{Text="v$($e.version) - $($e.type)"; Foreground="#777777"})
                       ) + (Create-GuiParamFields $e.param $e.id)
                    })
                }
            )}
            $GRID.Children.Add($border)
        }
        
        #-- add menu entry
        if (($e.type -notmatch "^init") -and ($e.category)) {
            $m = $GUI.w.findName("Menu_$($CAT)")
            # new Extras > submenu if not found
            if (-not $m) {
                $m = (W MenuItem @{Name=$CAT; Header=$e.category})
                $GUI.w.findName("Menu_EXTRAS").Items.Add($m)
                $GUI.w.registerName("Menu_$CAT", $m)
            }
            # add
            $ICON = W Image @{Source=(Get-IconPath $e.img $e.icon $e.category); Height=20; Width=20}
            $m.Items.Add((W MenuItem @{Header=$e.title; InputGestureText=$e.keycode; Icon=$ICON; ToolTip=(W ToolTip @{Content=$e.description}); Add_Click=$cb}))
        }
    }

    #-- and a shortcut - have their own custom sorting from `#shortcut: 123`
    if ($m = $GUI.w.findName("Shortcuts")) {
        ForEach ($e in ($menu | ? { $_.shortcut -match "\d+|true|yes" } | Sort-Object @{expression={$_.shortcut.trim()}} )) {
            $cb = { $GUI.tasks += $e }.getnewclosure()
            $ICON = W Image @{Source=(Get-IconPath $e.img $e.icon $e.category); Height=22; Width=22}
            $BTN = W Button @{Style=$GUI.styles.ActionButton; ToolTip=(W ToolTip @{Content=$e.title}); Add_click=$cb; Height=22; Width=22; Content=$ICON}
            $m.Children.Add($BTN)
        }
    }
}


#-- attach callbacks for main UI buttons
function Add-ButtonHooks {
    <#
      .SYNOPSIS
         Prepares callbacks for buttons/toolbar in main window.
      .DESCRIPTION
         Such as the computer and username fields, or the clipboard functionality.
      .NOTES
         Ping and username lookup can freeze the UI, as they run in the WPF runspace already.
         $GUI.machine and $GUI.username are shared with the main thread.
         As is the clipboards $GUI.html shadow content.
    #>

    #-- computer name
    $GUI.w.findName("BtnComputer").add_Click({ 
        if ($m = Get-Clipboard) {
            $GUI.machine.Text = $m
            $col = if (Test-Connection $m -Count 1 -Quiet -TTL 10 -ErrorAction SilentlyContinue) {"#5599ff99"} else {"#55ff7755"}
            $GUI.machine.Items.Insert(0, (W ComboBoxItem @{Content=$m; Background=$col}))
            $GUI.machine.Background = "$col"
        }
    })
    $GUI.w.findName("BtnComputerClr").add_Click({ $GUI.machine.Text = ""; $GUI.machine.Background = "White" })
    $GUI.w.findName("BtnComputerCpy").add_Click({ Set-Clipboard $GUI.machine.Text })
    $GUI.w.findName("BtnComputerPng").add_Click({ })
    $GUI.w.findName("BtnComputerUsr").add_Click({ if ($u = Get-MachineCurrentUser $GUI.machine.Text) { $GUI.username.Text = $u } })

    #-- user name
    $GUI.w.findName("BtnUsername").add_Click({
        $u = Get-Clipboard
        if ($u -match "\w+[@, ]+\w+") { $u = (AD-Search $u -only 1) }
        $GUI.username.Text = "$u"
    })
    $GUI.username.add_DropDownOpened({
        if (($u = $GUI.username.Text).length -gt 2) {
            $GUI.username.Items.Clear()
            AD-Search $u | % { $GUI.username.Items.Add($_) }
        }
    })
    $GUI.username.add_DropDownClosed({
        $GUI.username.Text = $GUI.username.Text -replace "\s*\|.+$", ""
    })
    $GUI.w.findName("BtnUsernameClr").add_Click({ $GUI.username.Text = "" })
    $GUI.w.findName("BtnUsernameCpy").add_Click({ Set-Clipboard $GUI.username.Text })
    $GUI.w.findName("BtnUsernameCom").add_Click({  })

    #-- bulk/csv
    $GUI.w.findName("BtnBulkimport").add_Click({ $GUI.bulkcsv.text = Get-NotepadCSV $GUI.bulkcsv.text "notepad" })

    #-- clipboard tools
    $GUI.w.findName("BtnClipText").add_Click({ Set-Clipboard $GUI.output.text })
    $GUI.w.findName("BtnClipHtml").add_Click({
        # use $html when available
        if ($GUI.html) {
            Set-ClipboardHtml $GUI.html
        }
        else {
            #@ToDo: convert TextBlock.Inlines to HTML
            Set-Clipboard ($GUI.output.text)
        }
    })
    $GUI.w.findName("BtnClipFree").add_Click({
        $prev_output = ($GUI.output.Inlines | % { $_ })
        $GUI.html = ""
        $GUI.output.text = ""
    })
    $GUI.w.findName("BtnClipSwap").add_Click({
        $GUI.output.text = ""
        $prev_output | % { $GUI.output.Inlines.Add($_) }
    })
    $global:prev_output = @()
    
    #-- window closing
    $GUI.w.add_Closed({
        $GUI.closed = $true
    })

    #-- unicode symbols
    if ($ls = $GUI.w.findName("UnicodeClip")) {
        $cb = {Set-Clipboard $this.Content}
        ForEach ($btn in $ls.Children) {
            $btn.Add_click($cb)
        }
    }
}

#-- User lookup (simple `finduser` variant)
function AD-Search {
    <#
      .SYNOPSIS
         Simple AD username search
      .DESCRIPTION
         Gets invoked whenever somethings is pasted into `username` input field.
         Defaults to search for "Last, First" names and spit out "UserName123".
         Also may scan for phone numbers or email addresses (used in UI for dropdown).
      .NOTES
         Defaults to [adsisearched], but may use Get-ADUser when available. By default
         the ActiveDirectory module is not loaded again into the WPF/UI Runspace.
         May require customization if AD displayname does not follow "First, Last" scheme.
    #>
    Param(
        $s = "",
        $only = 0,
        $cache = 0
    )
    $s = ($s -replace "%","*")
    #-- AD search
    if (Test-Path function:Get-ADUser) {
        $filter = switch -regex ($s) {
           "^[*?\d]+$" {   "telephoneNumber -like '*$s'"                              }
           "^\w+, \w+" {   "displayname -like '$s*'"                                  }
             "\w+@\w+" {   "mail -like '$s'"                                          }
               default {   "displayname -like '$s*' -or samaccountname -like '$s*'"   }
        }
        #-- find
        $u = Get-ADUser -Filter $filter -Properties samaccountname,displayname,telephonenumber
    }
    #-- crude cache search
    elseif ($cache) {
        return ((Get-Content "data\adsearch.txt") -match $s)
    }
    #-- ADSI query
    else {
        $s = $s -replace '([;\\"#+=<>])', '\$1'
        $filter = switch -regex ($s) {
           "^[*?\d]+$" {   "telephonenumber=*$s"                              }
        "^\w+\\?, \w+" {   "displayname=$s"                                   }
           ".+\w+@\w+" {   "mail=$s"                                          }
               default {   "|(displayname=$s*)(samaccountname=$s*)"           }
        }
        $adsi = [adsisearcher]"(&(objectClass=user)(objectCategory=person)($filter))"
        [void]$adsi.PropertiesToLoad.AddRange(("telephonenumber","displayname","samaccountname","mail","*","+"))
        $u = ($adsi.findAll() | % { $_.properties })
    }
    #-- result
    if ($only) { return [string]$u.samaccountname }
    else { return @($u | % { "{0} | {1} | {2}" -f [string]$_.samaccountname, [string]$_.displayname, [string]$_.telephonenumber }) }
}

#-- get current user
function Get-MachineCurrentUser($m) {
    <#
      .SYNOPSIS
         Get current user from remote machine name.
      .DESCRIPTION
         Runs a quick WMI query. Returns user name (domain prefix stripped).
      .NOTES
         Does not fallback to alternative Explorer.exe process scan.
    #>
    if ($m -and ($w = gwmi win32_computersystem -computername $m) -and ($u = $w.username)) {
        return $u -replace "^\w+[\\\\]",""
    }
}


#-- HTML output
#   · invoked by Out-Gui
function Out-Html {
    <#
      .SYNOPSIS
         Assembles HTML (clipboard) output from each Out-Gui (Write-Host alias) call.
      .DESCRIPTION
         Converts -F foreground and -B background colors and appends to $GUI.html.
         Additionally provides -Bold and -Underline support (but those aren't Write-Host/CLI compatible then).
      .NOTES
         That's the lazy approach. Original plan was to convert WPF TextBlock inlines in a HTML clipboard function.
         Postponed since Out-Gui does not implement any images or table output yet anyway.
    #>
    param(
        [Parameter(ValueFromPipeline=$true)]$str = $null,
        [alias("Foreground")]$f = $null,
        [alias("Background")]$b = $null,
        [alias("Bld")]$bold = $null,
        [alias("Strikethrough")]$S = $null,
        [alias("Underline")]$U = $null,
        [alias("NoNewLine")][switch]$N = $false,
        $css = ""
    )

    #-- foreground/background maps (for white HTML background)
    $map_f = @{
        Yellow = "#5554400"
        Red = "#55070F"
        Blue = "#111144"
        Green = "#004400"
        Cyan = "#004455"
        White = "#111"
        Gray = "#666"
        Black = "#222"
    }
    $map_b = @{
        Yellow = "#f3f3aa"
        Green = "#aaeeaa"
        Blue = "#a5a5f5"
        Cyan = "#a1ece5"
        Red = "#ef8a77"
        White = "#444"
        Gray = "#bbb"
        Black = "#555"
    }

    #-- normalize string
    #if (!$str) { return }
    $str = ($str | Out-String -Width 120).trim()
    $str = $str -replace "(?m)\s*$", ""
    $str = $str -replace "<","&lt;"
    $str = $str -replace ">","&gt;"
    $str = $str -replace "\r?\n", "<br>`r`n"

    #-- colorize
    $str = $str -replace "✔","<font color=#115511>✔</font>"
    $str = $str -replace "✘","<font color=#771122>✘</font>"
    if ($f -eq "#ff9988dd") {
        $f = $null;      # detect script title/description block
        $bold = "1";
        $b = "#eeeef5";
    }
    if ($f -and $map_f.containsKey($f)) {
        $f = $map_f[$f]       # remap to darker foreground
    }
    if ($b) {
        if ($map_b.containsKey($b)) {
            $b = $map_b[$b]   # lighter background colors
        }
        $css += "background-color:$b;"
    }
    if ($bold) {
        $str = "<strong>$str</strong>"
    }
    if ($U) {
        $str = "<u>$str</u>"
    }
    if ($css -or $f) {
        $str = "<font style='$css' color='$f'>$str</font>"
    }
    if ($N) { $NL = "" } else { $NL = "<br>`r`n" }
    $GUI.html += $str + $NL
}


#-- GUI interactions
#   · Primarily appends to `output` TextBlock
#   · Is called from main thread (and the only allowed interaction method),
#     itself uses w.Dispatcher.Invoke to execute codeblock in GUI runspace
#   · Also updates window `-Title`
#   · And provides `-GetVars` workaround
filter Out-Gui {
    <#
      .SYNOPSIS
         Fills the $GUI.output TextBlock from calls to Write-Host (aliased to Out-Gui).
      .DESCRIPTION
         Mainly converts any objects (HashTables/PSObjects) to strings, then appends them to main output pane.
         Since this is called from the main script Runspace, uses the WPF dispatcher to execute the appending
         in the WPF runspace (factually WPF does it in another separate thread).
      .PARAM str
         Input string or object
      .PARAM f
         Foreground color
      .PARAM b
         Background color
      .PARAM N
         NoNewLine
      .PARAM bold, s
         Not CLI/Write-Host compatible. Should not be used unless the plugin/script was meant for GUI usage only.
      .NOTES
         Also implements another few UI interactions, such as -title setting, or output -clear.
         For example -getvars returns a HashTable of toolbar and plugin/ToolBlock input fields.

         Ultimately this should handle error/exception highlighting and creating OUtGrid-like display
         for HashTables and Objects. Not implemented yet, as it doesn't even run on PS 5.0 yet.
         Also should be rewritten to Begin{} Process{} End{} and split up to handle proper piping.

         (Currently this does not allow to synchronously handle output from tools/scripts. Write-Host
         calls get processed before any Out-Default remains...)
    #>
    param(
        [Parameter(ValueFromPipeline=$true)]$str = $null,
        [alias("Foreground")]$f = $null,
        [alias("Background")]$b = $null,
        [alias("Bld")]$bold = $null,
        [alias("Strikethrough")]$S = $null,
        [alias("Underline")]$U = $null,
        [alias("NoNewLine")][switch]$N = $false,
       [string]$title = $null,
       [string]$getvars = $null,
       [string]$testpfx = $null,
       [switch]$clear = $false
    )

    trap { $_ | Out-String >> ./data/log.errors.wpf.txt }

    Out-Html $str -f $f -b $b -bold $bold -s $s -U $U -N:$N
    

    #-- wrap in scriptblock for GUI thread dispatcher
    $callback = [action]{
        $inlines = @()
        trap { $_ | Out-String >> ./data/log.errors.out-gui.txt }
    
        #-- clear stale contents
        #   · if last output over 5 minutes ago (`-clear:$true` is set by caller)
        if ($clear) {
            $GUI.prev_output = $GUI.output.inlines
            $GUI.output.text = ""
            $GUI.html = ""
        }

        #-- just -Title update
        if ($title) {
            $GUI.w.title = $title
            return
        }

        #-- Read input fields
        #  · this is a workaround, because accessing $GUI.machine.text
        #    directly from parent runspace would hang up WPF
        #  · for some reason also needs hashtable recreated
        if ($getvars) {
            $GUI.vars = @{}
            $getvars.split("[,;]") | % {
                if (($key = $_.trim()) -and ($f = $GUI.w.findName($key)) -or ($f = $GUI.w.findName("var_$($testpfx)_$($key)"))) { 
                    $GUI.vars.$key = $f.text
                }
            } | Out-Null
            return $GUI.vars
        }

        # see if ErrorRecord
        try {
            $type = $str.GetType().Name
        }
        catch {
            $type = "unknown"
        }
        if ($type -eq "--ErrorRecord") {
            $inlines += WD Figure @{Background="#ff773322"; Content=($str|Out-String)}
        }
        
        # or HashTable
        elseif ($type -eq "---HashTable") {
            $inlines += ($str|Out-String)
        }

        # or image
        elseif ($type -eq "Image" -or $type -eq "Paragraph" -or $type -eq "InlineUIContainer") {
            # leave as-is
        }

        # plain string
        else {
            $str = $str | Out-String -Width 100
            $a = WD Run @{Text=$str}
            if ($f) { $a.Foreground=$f }
            if ($b) { $a.Background=$b }
            if ($bold) { $a.FontWeight="Bold" }
            if ($S) { $a.TextDecorations="Strikethrough" }
            if ($U) { $a.TextDecorations="Underline" }
            $inlines += $a
        }

        #-- append
        $inlines | % { $GUI.output.Inlines.Add($_) } | Out-Null
        [void]$GUI.output.Parent.ScrollToBottom()
    }
    
    #-- PowerShell >= 3.0 does need the parameters swapped
    #   (@src: https://gallery.technet.microsoft.com/Script-New-ProgressBar-329abbfd)
    if ($PSVersionTable.PSVersion.Major -eq 2) {
        [void]$GUI.w.Dispatcher.Invoke("Normal", $callback)
    }
    else {
        [void]$GUI.w.Dispatcher.Invoke($callback)
    }
}


##########################################################################################
########################   everything below runs in main thread   ########################
##########################################################################################


#-- User input
#   · this is aliased over `Read-Host`
#   → so scripts can be used unaltered in standalone/CLI or GUI mode
#   · for standard field names just returns the GUI $machine/$username input
#   · else shows a VB input box window
function Ask-Gui {
    <#
      .SYNOPSIS
         Alias over `Read-Host`
      .DESCRIPTION
         Compares Read-Host requests against a list of know variable names/titles,
         such as "computer" or "username". If a match is found, returns the current
         value from the GUI input field. Else shows a VBS popup.

         The purpose is to allow tools/scripts to run in CLI mode as well as in GUI
         frontend unchanged. But to avoid needless popups for predefined input fields.
      .PARAM str
         Textual input query. Usually just the variable name '$computer' or '$user'.
         Aliases like 'Machine' and 'PC' and 'Hostname' are known, as well as 'AD-Account'
         or 'AD-User-Name' e.g.
      .NOTES
         Scripts should preferrably query for input in their Param() section once,
         and list custom fields per plugin meta #param: list.
         The GUI frontend does not implement a full PSHost, just an output TextBlock.
         (Input fields might be feasible, but too much work IMHO.)
    #>
    param($str, $title="Input", [switch]$N=$false)
    if ($str -match "^(?i)[$]?(AD[-\s]?)?(User|Account)(?:[-\s]?Name)?\s*[=:]*\s*?$") {
        return $GUI.vars.username
    }
    elseif ($str -match "^(?i)[$]?(Computer|Machine|PC|Host)([-\s]?Name)?\s*[=:]*\s*$") {
        return $GUI.vars.machine
    }
    elseif ($str -match "^(?i)[$]?(Bulk|bulk.?csv|bulkfn|list|CSV)") {
        return $GUI.vars.bulkcsv  # should be exported to filename?
    }
    elseif ($GUI.vars.containsKey($str)) {  #-- per-plugin input boxes
        return $GUI.vars[$str]
    }
    else {
        #-- Trivial input box for everything else
        return [Microsoft.VisualBasic.Interaction]::InputBox($title, $str, '')
    }
}

# alias for CMD `CHOICE` function
function Ask-GuiVar {
    param($P, $VAR, $M, $TEXT)
    $v = Ask-Gui $TEXT "Choice"
    Invoke-Expression ('$global'+":$VAR = '$v'")
}

# alias for CMD `CHOICE` function
function Ask-GuiMachine {
    return $global:machine
}

# alias for CMD `CHOICE` function
function Ask-GuiUser {
    return $global:username
}

#-- Converts `# param: name,list` from vars{} to quoted cmdline "arg" "strings"
function Get-ParamVarsCmd {
    <#
      .SYNOPSIS
         Crafts a list of cmd/Invoke-quoted strings from params list
      .DESCRIPTION
         Is used for type:window and type:cli scripts/plugins. Those get executed
         in a separate Powershell process, thus need input variables as exec arguments.
      .NOTES
         Somewhat tedious, as this is the third duplication of field/varname alias handling.
    #>
    Param(
        $param = "machine,username",         # from plugin meta field
        $vars = @{machine=""; username=""},  # from GUI -GetVars
        $out = ""
    )

    if (!$param) { $param = "machine,username" }   # default list
    ForEach ($key in $param.split("[,;]")) {
        if (!($key = $key.trim()).length) { continue }
        # aliases
        if ($key -match '^(host|hostname|computer|pc-?name)$') { $key = "machine" }
        if ($key -match '^(user|ad-?name|account|accountname)$') { $key = "username" }
        if ($key -match '^(bulk|bulkcsv|bulkfn|bulklist)$') {
            $tmpfn = [IO.Path]::GetTempFileName()
            $vars[$key] | Out-File $tmpfn -Encoding UTF8 | Out-Null
            $vars[$key] = $tmpfn   # ToDo: clean tmp via $plugins.after[]
        }
        # quote + append
        $out += ' "'+($vars[$key] -replace '([\\"^])','^$1')+'"'
    }
    return $out
}

#-- wrapper around Process-MenuTask
function Run-GuiTask {
    <#
      .SYNOPSIS
         Executes the given $e event queue entry returned from GUI (menu or button callback).
      .DESCRIPTION
         Checks the $menu.type and handles output start, script execution, and result collection.
          - Just prints .title and .description first
          - Then assembles GUI variables ($machine, $username and extra script fields)
          - Runs `type:inline` plugins in main thread.
          - But `type:window` in separate window/CLI process.
          - Captures and outputs any errors
          - Then returns tp the Start-Win loop
      .PARAM e
         The script/tool entry as returned from $GUI.tasks (just one of the $menu entries really).
         Said queue gets filled by any of the menu or toolblock buttons in the main window.
      .NOTES
         Ideally this should just wrap the CLI Process-MenuTask function.
         The | Out-Gui piping does not mix direct Write-Host calls in order yet. Thus any extra
         script output gets shown AFTERWARDS. (to be fixed with proper pipe handling)
    #>
    [CmdletBinding()]
    param($e, $clear=$false) # one of the $menu entries{}

    #-- check last output time (>= 5 minutes)
    if ($last_output -and (([double]::parse((Get-Date -u %s)) - $last_output) -gt $cfg.autoclear)) {
        $clear = $true
    }
    #-- print header
    Out-Gui -title "➱ KlickiBunti → $($e.title)" -clear:$clear
    if ((!$e.noheader) -and (!$cfg.noheader)) {
        Out-Gui -f '#ff9988dd' -b '#ff102070' ("$($e.title) - $($e.description)")
    }

    #-- Get-GuiVars (fetch input from "Ribbon"-fields: machine, username, ...)
    Out-Gui -GetVars "machine,username,bulkcsv,$($e.param)" -testpfx $e.id  # populates $GUI.vars{}
    $GUI.vars.GetEnumerator() | % { Set-Variable $_.name $_.value -Scope Global }
    
    #-- plugins
    TRAP { $_ | out-gui -b red }
    $plugins.before | % { Invoke-Expression ($_.ToString()) }

    #-- Run $menu entry rules (command=, func=, or fn=)
    try {
        #-- Internal commands
        if ($e.command) {
            [void]((Invoke-Expression $e.command) | Out-String | Out-Gui -f Yellow)
        }
        elseif ($e.func) {
            [void]((Invoke-Expression "$e.func $machine $username") | Out-String | Out-Gui -f Yellow)
        }
        #-- Start script
        elseif ($e.fn) {
            if ($e.type -match "window|cli") {  # in separate window
                $cmd_params = Get-ParamVarsCmd ($e.param) ($GUI.vars)
                Start-Process powershell.exe -Argumentlist "-STA -ExecutionPolicy ByPass -File $($e.fn) $cmd_params"
            }
            else {  # dot-source all "inline" type: plugins
                . $e.fn | Out-String | Out-Gui # -f Green
            }
        }
        #-- No handler
        else {
            Out-Gui -f Red "No Run-GuiTask handler for:"
            [void](@($e) | Out-String | Out-Gui)
        }
    }
    #-- Error cleanup
    catch {
        Out-Gui -f Yellow $_
    }
    if ($Error) {
        $Error | % { $_ | Out-Gui -f Red } | Out-Null
        [void]($Error.Clear())
    }

    #-- Run plugins / cleanup
    $plugins.after | % { Invoke-Expression ($_.ToString()) }
    $global:last_output = [double]::parse((Get-Date -u %s))
    Out-Gui -title "➱ KlickiBunti"
}


#-- verify runspace and window are still active
function Test-GuiRunning($shell, $GUI) {
    return ($shell.Runspace.RunspaceStateinfo.State -eq "Opened") -and (-not $GUI.closed)
}


#-- Initialize Window, Buttons, Subshell+GUI thread
function Start-Win {
    <#
      .SYNOPSIS
         Starts the WPF window and Runspace, then waits for GUI events
      .DESCRIPTION
         Sets up GUI environment from $menu entries, then polls the shared $GUI. variables
      .NOTES
         Perhaps the least interesting function and self-explanatory.
    #>
    Param($menu)

    #-- no threading / manual
    #    $GUI.w = $w = WPF-Window; $w.showdialog(); Write-Host "Ran WPF-Window without thread"

    #-- start GUI in separate thread/runspace
    $shell = New-GuiThread -Code {
        $GUI.w = $w = WPF-Window
        #TRAP { $_ | Out-Gui -f '#ffee99' -b '#cc5544'; $Error.Clear() }
        #TRAP { $GUI.w.Host.UI.Write($_);  }
        #Import-Module ActiveDirectory
        Add-ButtonHooks
        Add-GuiMenu $menu
        #-- execute `type:init` files right away
        $menu | ? { $_.fn -and ($_.type -eq "init-gui") } | % { . $_.fn }
        $w.ShowDialog()
    } -Vars "menu,BaseDir,cfg,plugins"
    $GUI.tasks = @()

    #-- wait for window to be visible
    while (!($GUI.w.IsInitialized)) { 
        Start-Sleep -milliseconds 275
        Write-Host 'wait on $GUI.w.isInitialized'
        #$global:GUI     |FT  |Out-String|Write-Host -f Blue
        #$shell.runspace   |FL -Prop *|Out-String|Write-Host -f Yellow
        #$shell.Streams    |FL|Out-String|Write-Host -f DarkGreen
        if ($Debug -and $shell.streams.error) {
            $shell.streams | FL | Out-String | Write-Host -b Red
        }
    } 

    #-- Alias console functions
    # `echo` is already alias to `Write-Output`
    Set-Alias -Name Write-Host  -Value Out-Gui  -Scope Global
    Set-Alias -Name Read-Host   -Value Ask-Gui  -Scope Global
    Set-Alias -Name Get-Machine -Value Ask-GuiMachine
    Set-Alias -Name choice      -Value Ask-GuiVar
    
    #-- basic error catching
    TRAP { $_ | Out-Gui -f Red; $Error.Clear() }

    #-- main loop / trivial message queue
    #   test for window state, pause a few seconds, then resolves $tasks[]
    while (Test-GuiRunning $shell $GUI) {
        if ($GUI.tasks) {
            $GUI.tasks | % { Run-GuiTask $_ }
            $GUI.tasks = @()
        }
        Start-Sleep -milliseconds 175
    }
}


