<#
.SYNOPSIS
    Experimental playground with unstable or hardcoded stuff.
#>

# Invocations
function open { & "c:\tools\totalcmd\TOTALCMD64.EXE" ($pwd) }

function edit( [string] $File, [switch] $SameEditor )
{
    $params = @()

    if( $SameEditor )
    {
        $params += "--reuse-window"
    }

    if( $File -match ":" )
    {
        $params += "--goto"
    }

    $params += $file

    # code --help | code -
    & code $params
}
