[CmdletBinding(DefaultParametersetName="p1")]
param(
    [Parameter(ValueFromPipeline=$true,Position=0)] [array] $ComputerName = $env:COMPUTERNAME,
    [Parameter(ValueFromPipeline=$true,Position=1,ParameterSetName="p1")] [string] $Title,
    [Parameter(ValueFromPipeline=$true,Position=2,ParameterSetName="p1")] [string] $GUID,
    [Parameter(ValueFromPipeline=$true,Position=1,ParameterSetName="p1")] [string] $Publisher,
    [switch] $List = $False
)

Begin {
}

Process{
    
    <#If ((Get-Service -ComputerName $ComputerName -DisplayName 'Remote Registry').Status -eq 'stopped') {
        $WasStopped = $true

        Write-Output "Enabling Remote Registry service..."
        (Get-WmiObject -class win32_service -ComputerName $ComputerName | where-object name -like 'remoteregistry').changestartmode("Automatic")

        Write-Output "Starting Remote Registry service..."
        Start-Service -DisplayName 'Remote Registry'
    } #>

    If (!$Title -and !$GUID) {
        $Title = "*"
    } ElseIf (!$Title) {
        $Title = $GUID
    }
    $Results = @()

    Function Get-Titles {
        param(
            [string] $Computer = $env:COMPUTERNAME,
            [string] $Title,
            [switch] $List = $False,
            [string] $Publisher
        )

            $SomethingsInstalled = $false
 
            #Write-Verbose $Publisher
            #Write-Verbose $Computer
 
            $ErrorActionPreference = "Continue"
            
            If ((Get-WmiObject win32_processor -ComputerName $Computer).AddressWidth -eq 32){   
                Write-Verbose -Verbose "[$computer] 32-bit system detected."
                $keys = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            } Else {   
                Write-Verbose -Verbose "[$computer] 64-bit system detected, iterating through 32-bit and 64-bit reg keys..."
                $keys = "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
            }
            
            Do {
                $Reg = ([microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $Computer))
                $AppList = $null
                ForEach ($key in $keys) {
                    ForEach ($regitem in $key) {
                        $RegKey = $Reg.OpenSubKey($regitem)
                        $SubKeys = $RegKey.GetSubKeyNames()
                        $Count = 0

                        $Data = ForEach ($skey in $SubKeys){   
                            $thisKey = $regitem + "\\" + $skey
                            $thisSubKey = $Reg.OpenSubKey($thisKey)
                            #write-output $thisSubKey.GetValue("DisplayName")
                            If (($thisSubKey.GetValue("DisplayName") -like $Title)){
                                
                                $SomethingsInstalled = $true    
                                $ErrorActionPreference = "SilentlyContinue"
                            
                                New-Object PSObject -Property @{
                                    Query = "$Title"
                                    Installed = $true
                                    InstallDate = $thisSubKey.GetValue("InstallDate")
                                    ComputerName = $Computer
                                    UninstallString = $thisSubKey.GetValue("UninstallString")
                                    Title = $thisSubKey.GetValue("DisplayName")
                                    Publisher = $thisSubKey.GetValue("Publisher")
                                    DisplayVersion = $thisSubKey.GetValue("DisplayVersion")
                                    InstallLocation = $thisSubKey.GetValue("InstallLocation")
                                    GUID = $($thisSubKey.GetValue("UninstallString")).Split("{}")[1]
                                }
                            } 
                         }

                        $Data
                        break

                        $Results += $Data

                    } #end for each regitem in keys
                } #end for each key in keys

                    If ($SomethingsInstalled -eq $false){
                        $Installed = New-Object PSObject -Property @{
                            Query = $Title
                            Installed = $false
                            InstallDate = 'Not installed'
                            ComputerName = $Computer
                            UninstallString = $null
                            DisplayName = $null
                            Publisher = $null
                            Title = $null
                            InstallLocation = $null
                            GUID = $null
                        }
                        $Results += $Installed
                    }
            } 
        
            Until ($Return -eq $null)

    $Results

    }
    
    ForEach ($Computer in $ComputerName){
            Get-Titles -Computer $Computer -Title $Title 
            #Start-Job -Name "$Computer job" -ScriptBlock $Scriptblock -ArgumentList $Computer,$Title,$List,$Publisher
    }

}

End {
    <#If ($WasStopped -eq $true) { 
        Write-Output "Stopping service 'Remote Registry'"
        Stop-Service -DisplayName 'Remote Registry' 
        Write-Output "Setting Startup to Manual for Remote Registry service"
        (Get-WmiObject -class win32_service -ComputerName $ComputerName | where-object name -like 'remoteregistry').changestartmode("manual")

    } #>

    #$Data | Select-Object * -Unique
}