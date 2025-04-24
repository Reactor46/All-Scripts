[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
param()

# Default console color setup
Complete-Once "Fonts" {
    Set-DefaultPowershellColors ".\%SystemRoot%_System32_WindowsPowerShell_v1.0_powershell.exe"
    Set-DefaultPowershellColors ".\%SystemRoot%_SysWOW64_WindowsPowerShell_v1.0_powershell.exe"
}
tm fonts

Complete-Once "ColorTool" {
    Push-Location "$PsScriptRoot\..\Bin\ColorTool\"
    .\ColorTool.exe -b -q campbell | Out-Null
    Pop-Location
}
tm ColorTool

# Cloud folders setup
switch ($env:ComputerName)
{
    "ALEXKO-LS"
    {
        $azcompute = "D:\src\mv\"
        $apgold = "D:\src\golds\ap\"
        $ntp = "D:\src\ntp\"
    }
    "ALEXKO-SB2"
    {
        $azcompute = "c:\src\mv\"
        $apgold = "c:\src\gold\ap\"
        $ntp = "C:\src\ntp\"
    }
}
tm "Variables setup"

# Set up environment variables
Set-EnvironmentVariable "AzCompute" $azcompute
Set-EnvironmentVariable "ApGold" $apgold
Set-EnvironmentVariable "NTP" $ntp
Set-EnvironmentVariable "Startup" "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
Set-EnvironmentVariable "Home" $env:USERPROFILE
tm "Environment setup"

# Tools folder creation
if( -not (Test-Path "c:\tools") )
{
    # Do we need to test that we have admin rights / that current user can do anything to drive c:\ ?
    mkdir "c:\tools" -ea Stop | Out-Null
}
tm "Tools root setup"

# Tools junction creation
foreach( $tool in ls $oneDrive\Distrib\tools -Directory -ea Ignore | where Name -notmatch "^_" )
{
    New-Junction $tool.FullName "c:\tools\$($tool.Name)"
}
tm "Tool junctions creation"

foreach( $tool in ls $oneDriveMicrosoft\tools -Directory -ea Ignore | where Name -notmatch "^_" )
{
    New-Junction $tool.FullName "c:\tools\$($tool.Name)"
}
tm "Microsoft specific tool junctions creation"

# The rest of the commands are possible only from an elevated prompt
if( -not (Test-Elevated) )
{
    return
}
tm "Elevation test"

# Shortcut creation
Copy-UpdatedFile "$PsScriptRoot\..\Shortcuts\Windows PowerShell.lnk" "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\System Tools\Windows PowerShell.lnk"
Copy-UpdatedFile "$PsScriptRoot\..\Shortcuts\Windows PowerShell.lnk" "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Windows PowerShell"
Copy-UpdatedFile "$PsScriptRoot\..\Shortcuts\LINQPad.lnk" "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\LINQPad.lnk"
tm "Shortcut creation"

# Common root junctions
New-Junction "c:\Program Files" "c:\Program Files (x86)\_x64_"
New-Junction "c:\Program Files (x86)" "c:\programs"
New-Junction $home "c:\home"
tm "Common junctions setup"

# Folder hide
"c:\Intel", "c:\PerfLogs", "c:\Program Files", "c:\Program Files (x86)", "c:\Users", "c:\Windows", "c:\inetpub" | where{gi $psitem -ea ignore} | Set-Visible $false
"$home\3D Objects", "$home\Contacts", "$home\Favorites", "$home\Links", "$home\Pictures", "$home\Saved Games", "$home\Searches" , "$home\Videos" | where{gi $psitem -ea ignore} | Set-Visible $false
tm "Folder hiding junctions setup"