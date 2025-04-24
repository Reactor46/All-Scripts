[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '')]
param()

# Elevated test
$SCRIPT:isElevated = Test-Elevated

# Don't use history in PsReadLine, it is buggy in the latest PS release
Set-PSReadlineOption -HistorySaveStyle SaveNothing

# History folder and file
$SCRIPT:historyFolder = "c:\automation\history\"
if( -not (Test-Path $historyFolder) )
{
    mkdir $historyFolder -ea Ignore | Out-Null
}
$SCRIPT:historyFile = "$historyFolder\{0}--$pid.ps1" -f [DateTime]::Now.ToString("yyyy.MM.dd--HH.mm.ss--UTCz")
$SCRIPT:lastCommandId = -1

# Prompt
function prompt
{
    $realLastExitCode = $LASTEXITCODE


    # Preserve last command in log file
    $lastCommand = Get-History -Count 1
    if( $lastCommand )
    {
        if( $lastCommand.Id -ne $SCRIPT:lastCommandId )
        {
            $SCRIPT:lastCommandId = $lastCommand.Id
            Add-Content $historyFile $lastCommand.CommandLine.Replace("`n", "`r`n")
        }
    }
    else
    {
        $lastCommandId = -1
        Add-Content $historyFile "# No history at this point of time"
    }

    # Update title
    if( $ExecutionContext.SessionState.LanguageMode -eq "FullLanguage" )
    {
        if( $GLOBAL:WindowTitle )
        {
            $host.UI.RawUI.WindowTitle = $GLOBAL:WindowTitle
        }
        else
        {
            $title = $pwd
            if( $env:home ) { $title = $title -replace [regex]::Escape($env:home), "~" }
            if( $env:inetroot )
            {
                $title = $title -replace [regex]::Escape($env:inetroot)
                if( -not $title ) { $title = "\" }
                $title = $title -replace [regex]::Escape("\src\Client\NTP"), "NTP"
                if( -not $title ) { $title = "NTP\" }
            }
            $host.UI.RawUI.WindowTitle = $title
        }
    }

    # Update prompt
    Write-Host "$pwd" -ForegroundColor DarkYellow -NoNewline
    Write-Host " [$Env:ComputerName] " -ForegroundColor DarkGreen -NoNewline
    if( $SCRIPT:isElevated )
    {
        Write-Host " ELEVATED" -ForegroundColor DarkCyan -NoNewline
    }
    Write-Host ""

    $LASTEXITCODE = $realLastExitCode
    [char] 187 + " "
}
