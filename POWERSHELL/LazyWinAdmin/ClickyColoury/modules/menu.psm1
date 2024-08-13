# api: ps
# type: functions
# title: Utility code
# description: Output and input shortcuts, screen setup, menu handling, meta extraction
# doc: http://fossil.include-once.org/clickycoloury/wiki/plugin+meta+data
# version: 0.8.4
# license: PD
# category: misc
# author: mario
# config:
#    { name: cfg.hidden, type: bool, value: 0, description: also show hidden menu entries in CLI version }
# status: beta
# priority: core
#
# A couple of utility funcs:
#  · Get-Machine        → Read-Host "machine"
#  · echo_n             → echo wo/ CRLF
#  · Edit-Config        → edit config file
#  · Extrac-PluginMeta  → tokenize comment fields
#  · preg_match_all     → PHP-esque regex matching
#  · Init-Screen        → print script summary
#  · Print-Menu         → output $menu 
#  · Print-MenuHelp     → show help= entries
#  · Process-Menu       → input/run $menu prompt
#
# Load with:
#  Import-Module. ".\modules\menu.psm1"


#-- get $machine
function Get-Machine() {
    [CmdletBinding()]
    Param($current = $global:machine)
    # Ask for '$machine'
    Write-Host -N -f Yellow '$machine'
    # Add (default/last)
    if ($current) {
        Write-Host -N -f Gray "($current)"
    }
    Write-Host -N -f Yellow ': '
    $new = (Read-Host).trim().toUpper()
    if (!$new) {
        return $current
    }
    # Ping test
    if (!(Test-Connection -ComputerName $new -Count 1 -BufferSize 128 -Quiet -ErrorAction SilentlyContinue)) {
        Write-Host -f Red " → not online"
    }
    $global:machine = $new
    return $new
}

#-- opens notepad for editing a list/csv file in-place
function Get-NotepadCSV() {
    Param($text="", $EDITOR=$env:EDITOR)
    $tmpfn = [IO.Path]::GetTempFileName()
    $text | Out-File $tmpfn -Encoding UTF8
    [void](Start-Process $EDITOR $tmpfn -Wait)
    $text = Get-Content $tmpfn | Out-String
    [void](Remove-Item $tmpfn)
    return $text
}

#-- echo sans newline
function echo_n($str) {
    [void](Write-Host -NoNewLine $str)
}

#-- start notepad on config file (obsolete, now in separate script)
function Edit-Config() {
    Param($fn, $options, $overwrite=0, $CRLF="`r`n", $EDITOR=$ENV:EDITOR)
    if ($overwrite -or !(Test-Path $fn)) {
        # create parent dir
        $dir = Split-Path $fn -parent
        if ($dir -and !(Test-Path $dir)) { md -Force "$dir" }
        # assemble defaults
        $out = "# type: config$CRLF# fn: $fn$CRLF$CRLF"
        $options | % {
            $v = $_.value
            switch ($_.type) {
                bool { $v = @('$False', '$True')[[Int32]$v] }
                default { $v = "'$v'" }
            }
            $out += '$' + $_.name + " = " + $v + "; # $($_.description)$CRLF"
        }
        # write
        $out | Out-File $fn -Encoding UTF8
    }
    & $EDITOR "$fn"
}

#-- Regex/Select-String -Allmatches as nested array (convenience shortcut)
function preg_match_all() {
    Param($rx, $str)
    $str | Select-String $rx -AllMatches | % { $_.Matches } | % { ,($_.Groups | % { $_.Value }) }
}

#-- baseline plugin meta data support
function Extract-PluginMeta() {
    <#
      .SYNOPSIS
         Reads top-comment block and plugin meta data from given filename
      .DESCRIPTION
         Plugin meta data is a cross-language documentation scheme to manage
         application-level feature plugins. This function reads the leading
         comment and convers key:value entries into a hash. Also prepares the
         config{} parameter list.
      .PARAMERER fn
         Script to read from
      .OUTPUTS
         Returns a HashTable of field: values, including the config: list/hash.
      .EXAMPLE
         In ClickyColoury ir reads the "plugins" at once like this:
           $menu = (Get-Iten tools*/*.ps1 | % { Extract-PluginMeta $_ })
         Entries than can be accessed like:
           $menu | % { $_.title -and $_.category -eq "network" }
      .NOTES
         Each entry contains an .id basename and .fn field, additionaly to what
         the plugin itself defines.
         Packs comment remainder as .doc field.
    #>
    Param($fn, $meta=@{}, $cfg=@())

    # read file
    $str = Get-Content $fn | Out-String

    # look for first comment block
    if ($m = [regex]::match($str, '(?m)((?:^\s*[#]+.*$\n?)+)')) {

        # remove leading #␣ from lines, then split remainder comment
        $str = $m.groups[1] -replace "(?m)^\s*#[ \t]*", ""
        $str, $doc = [regex]::split($str, '\r?\n\r?\n')

        # find all `key:value` pairs
        preg_match_all -rx "(?m)^([\w-]+):\s*(.*(?:$)(?:\r?\n(?!\w+:).+$)*)" -str $str | % { $meta[$_[1]] = $_[2].trim() }

        # split out config: and crude-parse it (JSONish serialization)
        preg_match_all -rx "\{(.+?)\}" -str $meta.config | % { $r = @{};
            preg_match_all -rx "([\w.-]+)\s*[:=]\s*(?:[']?([^,;}]+)[']?)" -str $_[1] | % {  $r[$_[1]] = $_[2] }; $cfg += $r;
        }

        # merge into hashtable
        $meta.fn = "$fn"
        $meta.id = ($fn -replace "^.+[\\/]|\.\w+$","") -replace "[^\w]","_"
        $meta.doc = ($doc -join "`r`n")
        $meta.config = $cfg
    }
    return $meta  # or return as (New-Object PSCustomObject -Prop $meta)
}


#-- script header
function Init-Screen() {
    param($x=80,$y=45)
    #-- screen size better than `mode con` as it retains scrolling:
    if ($host.Name -match "Console") {
        $con = $host.UI.RawUI
        $buf = $con.BufferSize; $buf.height = 16*$y; $buf.width = $x; $con.BufferSize = $buf;
        $win = $con.WindowSize; $win.height =    $y; $win.width = $x; $con.WindowSize = $win;
    }
    #-- header
    $meta = $cfg.main
    Write-Host -b DarkBlue -f White ("  {0,-60} {1,15} " -f $meta.title, $meta.version)
    Write-Host -b DarkBlue -f Gray  ("  {0,-60} {1,15} " -f $meta.description, $meta.category)
    try { $host.UI.RawUI.WindowTitle = $meta.title } catch { }
}

#-- group plugin list by category, sort by sort: / key: / title:
function Sort-Menu($menu) {
    $usort_cat = { (@{cmd=1; powershell=2; onbehalf=3; exchange=4; empirum=5; network=6; info=7; wmi=8}[$_.category], $_.category) -ne $null }
    $usort_key = { if ($_.key -match "(\d+)") { [int]$matches[1] } else { $_.key } }
    return ($menu | Sort-Object $usort_cat, {$_.sort}, $usort_key, {$_.title})
}

#-- string cutting
function substr($str, $from, $to) {
    if ($to -lt $str.length) {
        $str.substring($from, $to)
    }
    else {
        $str
    }
}

#-- Write out menu list (sorted, 3 columns, with category headers)
function Print-Menu() {
    param($menu, $cat=".+", $last_cat="", $i=0)
    # group by category
    $ls = Sort-Menu ($menu | ? { $_.title -and $_.key -and ($_.category -match $cat) -and ((!$cfg.hidden) -or !$_.hidden) })
    $ls | % {
        if ($last_cat -ne $_.category) {
            if ($line) { Write-Host ""}
            Write-Host -f Black ("     {0,-74}" -f ($last_cat = $_.category))
            $i = 0
        }
        $line = (($i++) % 3 -ne 2)
        Write-Host -N -f Green ("{0,4}" -f (substr $_.key.split("|")[0] 0 4))
        Write-Host -N -f DarkRed  ("→")
        Write-Host -N:$line -f White ("{0,-21}" -f (substr $_.title 0 21))
    }
    echo ""
}

#-- print help= entries from $menu (→ not very pretty)
function Print-MenuHelp($menu) {
    $menu | ? { $_.title -and $_.key } | % {
        Write-Host -N -f Green (" " + $_.key.split("|")[0..2] -join ", ").PadRight(15)
        Write-Host -f White (" " + $_.title)
        Write-Host -f Gray ("                " + ($_.description))
    }
}

#-- Invoked on one menu entry → executes .command, or .func, or loads .fn script
filter Process-MenuTask() {
    Param($params)
    $host.UI.RawUI.WindowTitle = "MultiTool → $($_.title)"
    Write-Host -b DarkBlue -f Cyan ("{0,-60} {1,18}" -f $_.title, $_.version)
    echo ""
    #-- commands or function
    try {
        if ($_.command) {
            Invoke-Expression $_.command  # no options
        }
        elseif ($_.func) {
            Invoke-Expression "$($_.func) $($params)" # pass optional flags
        }
        #-- no fn?
        elseif (!$_.fn) {
            Write-Host -f Red "No processor for task:"
            $_ | FT | Write-Host
        }
        #-- start in separate "window"
        elseif ($_.type -eq "window") {
            Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy ByPass -File $($_.fn) $global:machine"
        }
        #-- dot-source file simply (for e.g. "inline" and "cli" type)
        else {
            Invoke-Expression ". $($_.fn) $($params)"  # run script
        }
    }
    catch {
        Write-Host -b DarkRed -f White ($_ | Out-String)
        $Error.Clear()
    }
    $host.UI.RawUI.WindowTitle = "MultiTool"
}

#-- Promp/input loop (REPL)
function Process-Menu() {
    param($menu, $prompt="Func")

    #-- prompt+exec loop
    while ($True) {
        Write-Host -N -f Yellow "$prompt> "
        $which = (Read-Host).trim()
        if ($which -match "^([\w.-]+)(?:\b|\s|$)(.*)$") { $params = $matches[2] } else { $params = $null }

        # find according menu entry: run func or expression
        $menu  |  ? { $_.key -and $which -match "^($($_['key']))\b" }  |  % { 
            while ($true) {
                $_ | Process-MenuTask $params
                if ((!$_.repeat) -or ((Read-Host "Wiederholen?") -notmatch "^[jy1rw]")) { break; }
            }
        }
    }
}

