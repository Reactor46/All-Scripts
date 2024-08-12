<#  
 
.SYNOPSIS  
    Run a report to find who is in Outlook Cache Mode
 
.NOTES  
    File Name  : OutlookCachingReport.ps1  
    Author     : Jose Espitia
    Requires   : PowerShell V5
    Version    : Version 1.00
 
#>
 
# Specify where to save the report
$Directory = "C:\LazyWinAdmin\Exchange\Logs"
$File = "MyReport.csv"
 
# Computer list
#Get-ADComputer -SearchBase "OU=IT Infrastructure,OU=Computers,OU=IT,OU=Las_Vegas,DC=contoso,DC=com" -Filter {Operatingsystem -Like 'Windows 10*' -and Enabled -eq 'true'}  -ErrorAction SilentlyContinue -Properties *  |
#    Select -ExpandProperty Name | Out-File -FilePath C:\LazyWinAdmin\Exchange\Clients.txt -Append
#
#Get-ADComputer -SearchBase "OU=IT Infrastructure,OU=Computers,OU=IT,OU=Las_Vegas,DC=contoso,DC=com" -Filter {Operatingsystem -Like 'Windows 7*' -and Enabled -eq 'true'} -ErrorAction SilentlyContinue -Properties * |
#    Select -ExpandProperty Name | Out-File -FilePath C:\LazyWinAdmin\Exchange\Clients.txt -Append

$ComputerList = "C:\LazyWinAdmin\Exchange\Clients.txt"
$Computers = Get-Content "$ComputerList"
 
ForEach($Computer in $Computers) {
     
    Try { 
        # Test connection with computer
        Test-Connection -ComputerName $Computer -ErrorAction Stop -Quiet
        # Query remote machines
        $HKEY_Users = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)
        # Get list of SIDs
        $SIDs = $HKEY_Users.GetSubKeyNames() | Where-Object { ($_ -like "S-1-5-21*") -and ($_ -notlike "*_Classes") }
 
        # Associate SID with Username
        $TotalSIDs = ForEach ($SID in $SIDS) {
            Try {
                $SID = [system.security.principal.securityidentIfier]$SID
                $user = $SID.Translate([System.Security.Principal.NTAccount])
                New-Object PSObject -Property @{
                    Name = $User.value
                    SID = $SID.value
                }                 
            } 
            Catch {
                Write-Warning ("Unable to translate {0}.`n{1}" -f $UserName,$_.Exception.Message)
            }
        }
        $UserList = $TotalSIDs
 
        # Loop through users to determine If they are in cache mode
        ForEach($User in $UserList) {
            # Get SID
            $UserSID = $User.SID
     
            # Get list of Outlook profiles
            $OutlookProfiles = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook\")
     
            # Loop through Outlook profiles to find caching key
            ForEach($Profile in ($OutlookProfiles.GetSubKeyNames())) {
         
                $ProfileKey = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook\$Profile")
         
                # Locate cache key
                If(($ProfileKey.GetValueNames() -contains "00036601") -eq $True) {
                    $Result = $ProfileKey.GetValue("00036601")
                    # Convert value to HEX
                    $Result = [System.BitConverter]::ToString($Result)
             
                    # Determine if cache mode is enabled
                    If($Result -like "8*") {
                        $CacheMode = "Enabled"
                    }
                    Else {
                        $CacheMode = "Disabled"
                    }
                    # Create custom table
                    $Table = New-Object PSObject -Property @{
                        Username = $User.Name
                        SID = $User.SID
                        "Computer Name" = $Computer
                        "Cache Mode" = $CacheMode
                        "Registry Key Value" = $Result
 
                    } | Select-Object Username, SID, "Computer Name", "Cache mode", "Registry Key Value"
                    # Export table to CSV
                    $Table | Export-Csv -NoTypeInformation -Append -Path "$directory\$file"
                }
             
            }
        }
    }
    Catch {
        # Create custom table
        $Table = New-Object PSObject -Property @{
            Username = "N/A"
            SID = "N/A"
            "Computer Name" = $Computer
            "Cache Mode" = "N/A"
            "Registry Key Value" = "N/A"
 
        } | Select-Object Username, SID, "Computer Name", "Cache mode", "Registry Key Value"
        # Export table to CSV
        $Table | Export-Csv -NoTypeInformation -Append -Path "$directory\$file"
    }
}