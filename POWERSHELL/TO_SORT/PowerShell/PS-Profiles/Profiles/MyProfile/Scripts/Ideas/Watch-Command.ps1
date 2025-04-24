##
## Author   : Roman Kuzmin
## Synopsis : Watch one screen of commands output repeatedly
## Modified : 2006.11.06
##
## -Command: script block which output is being watched
## -Seconds: refresh rate in seconds
##
## *) Commands should not operate on console.
## *) Tabs are simply replaced with spaces.
## *) * indicates truncated lines.
## *) Empty lines are removed.
## *) Ctrl-C to stop.
##

<#
# PowerShell commands:
Watch-Command {Get-Process}
Watch-Command {Get-Process | Format-Table Name, Id, WS, CPU, Description, Path -Auto}

# Native executable commands:
Watch-Command {netstat -e -s}
Watch-Command {cmd /c dir /a-d /o-d %temp%}
#>

param([scriptblock]$Command_ = {Get-Process}, [int]$Seconds_ = 3)

$private:sb = New-Object System.Text.StringBuilder
$private:w0 = $private:h0 = 0
for(;;) {

    # invoke command, format output data
    $private:n = $sb.Length = 0
    $private:w = $Host.UI.RawUI.BufferSize.Width
    $private:h = $Host.UI.RawUI.WindowSize.Height-1
    [void]$sb.EnsureCapacity($w*$h)
    .{
        & $Command_ | Out-String -Stream | .{process{
            if ($_ -and ++$n -le $h) {
                $_ = $_.Replace("`t", ' ')
                if ($_.Length -gt $w) {
                    [void]$sb.Append($_.Substring(0, $w-1) + '*')
                }
                else {
                    [void]$sb.Append($_.PadRight($w))
                }
            }
        }}
    }>$null

    # fill screen
    if ($w0 -ne $w -or $h0 -ne $h) {
        $w0 = $w; $h0 = $h
        Clear-Host; $private:origin = $Host.UI.RawUI.CursorPosition
    }
    else {
        $Host.UI.RawUI.CursorPosition = $origin
    }
    Write-Host $sb -NoNewLine
    $private:cursor = $Host.UI.RawUI.CursorPosition
    if ($n -lt $h) {
        Write-Host (' '*($w*($h-$n)+1)) -NoNewLine
    }
    elseif($n -gt $h) {
        Write-Host '*' -NoNewLine
    }
    $Host.UI.RawUI.CursorPosition = $cursor
    Start-Sleep $Seconds_
}

