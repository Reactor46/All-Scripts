## Begin Test-RegistryKey
Function Test-RegistryKey {
    <#
    .SYNOPSIS
    Tests if a registry Key and/or value exists. Can also test if the value is of a certain type.
    .DESCRIPTION
    Provides a reliable way to check registry values even when they are empty or null.
    .PARAMETER Key
    Path of the registry key (Required).
    .PARAMETER Name
    The value name (optional).
    .PARAMETER Value
    Value to compare against (optional).
    .PARAMETER Type
    The expected type of the registry value. Options: 'Binary', 'DWord', 'MultiString', 'QWord', 'String'.
    .EXAMPLE
    Test-RegistryKey -Key "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Install
    # Returns $true if the value exists.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory=$false)]
        [string]$Name,

        [Parameter(Mandatory=$false)]
        $Value,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Binary','DWord','MultiString','QWord','String')]
        [Microsoft.Win32.RegistryValueKind]$Type
    )

    Begin {
        $CmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
        Write-Verbose "Running function: $CmdletName"
    }

    Process {
        $Key = Convert-RegistryPath -Key $Key

        if (-not (Test-Path -Path $Key -PathType Container)) {
            Write-Verbose "Key [$Key] does NOT exist."
            return $false
        }

        $AllKeyValues = Get-ItemProperty -Path $Key

        if ($Name) {
            if (-not ($AllKeyValues.PSObject.Properties.Name -contains $Name)) {
                Write-Verbose "Key [$Key] exists, but value [$Name] does not."
                return $false
            }

            $ValDataRead = $AllKeyValues.$Name
            Write-Verbose "Value [$Name] exists with data [$ValDataRead]."

            if ($Value -and $Value -eq $ValDataRead) {
                return $true
            }

            if ($Type) {
                $ActualType = switch ($ValDataRead.GetType().Name) {
                    "String"   { 'String' }
                    "Int32"    { 'DWord' }
                    "Int64"    { 'QWord' }
                    "String[]" { 'MultiString' }
                    "Byte[]"   { 'Binary' }
                    default    { 'Unknown' }
                }

                Write-Verbose "Value [$Name] is of type [$ActualType]."
                return ($ActualType -eq $Type)
            }

            return $true
        }

        Write-Verbose "Key [$Key] exists."
        return $true
    }

    End {
        Write-Verbose "Function [$CmdletName] execution completed."
    }
}
## End Test-RegistryKey
## Begin Set-RemoteRegistryKey
Function Set-RemoteRegistryKey {
    <#
    .SYNOPSIS
        Set registry key on remote computers.
    .DESCRIPTION
        This Function uses .Net class [Microsoft.Win32.RegistryKey].
    .PARAMETER ComputerName
        Name of the remote computers.
    .PARAMETER Hive
        Hive where the key is.
    .PARAMETER KeyPath
        Path of the key.
    .PARAMETER Name
        Name of the key setting.
    .PARAMETER Type
        Type of the key setting.
    .PARAMETER Value
        Value tu put in the key setting.
    .EXAMPLE
        Set-RemoteRegistryKey -ComputerName $env:ComputerName -Hive "LocalMachine" -KeyPath "software\corporate\master\Test" -Name "TestName" -Type String -Value "TestValue" -Verbose
    .LINK
        http://itfordummies.net
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage='Provide a ComputerName')]
        [String[]]$ComputerName=$env:ComputerName,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]$Hive,
        
        [Parameter(Mandatory=$true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory=$true)]
        [String]$Name,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryValueKind]$Type,
        
        [Parameter(Mandatory=$true)]
        [Object]$Value
    )
    Begin{
    }
    Process{
        ForEach ($Computer in $ComputerName) {
            try {
                Write-Verbose "Trying computer $Computer"
                $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", "$Computer")
                Write-Debug -Message "Contenur de Reg $reg"
                $key=$reg.OpenSubKey("$KeyPath",$true)
                if($key -eq $null){
                    Write-Verbose -Message "Key not found."
                    Write-Verbose -Message "Calculating parent and child paths..."
                    $parent = Split-Path -Path $KeyPath -Parent
                    $child = Split-Path -Path $KeyPath -Leaf
                    Write-Verbose -Message "Creating the subkey $child in $parent..."
                    $Key=$reg.OpenSubKey("$parent",$true)
                    $Key.CreateSubKey("$child") | Out-Null
                    Write-Verbose -Message "Setting $value in $KeyPath"
                    $key=$reg.OpenSubKey("$KeyPath",$true)
                    $key.SetValue($Name,$Value,$Type)
                }
                else{
                    Write-Verbose "Key found, setting $Value in $KeyPath..."
                    $key.SetValue($Name,$Value,$Type)
                }
                Write-Verbose "$Computer done."
            }#End Try
            catch {Write-Warning "$Computer : $_"} 
        }#End ForEach
    }#End Process
    End{
    }
}
## End Set-RemoteRegistryKey
## Begin Search-Registry
Function Search-Registry {
<#
.SYNOPSIS
Searches registry key names, value names, and value data (limited).

.DESCRIPTION
This Function can search registry key names, value names, and value data (in a limited fashion). It outputs custom objects that contain the key and the first match type (KeyName, ValueName, or ValueData).

.EXAMPLE
Search-Registry -Path HKLM:\SYSTEM\CurrentControlSet\Services\* -SearchRegex "svchost" -ValueData

.EXAMPLE
Search-Registry -Path HKLM:\SOFTWARE\Microsoft -Recurse -ValueNameRegex "ValueName1|ValueName2" -ValueDataRegex "ValueData" -KeyNameRegex "KeyNameToFind1|KeyNameToFind2"

#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0, ValueFromPipelineByPropertyName)]
        [Alias("PsPath")]
        # Registry path to search
        [string[]] $Path,
        # Specifies whether or not all subkeys should also be searched
        [switch] $Recurse,
        [Parameter(ParameterSetName="SingleSearchString", Mandatory)]
        # A regular expression that will be checked against key names, value names, and value data (depending on the specified switches)
        [string] $SearchRegex,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that key names will be tested (if none of the three switches are used, keys will be tested)
        [switch] $KeyName,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that the value names will be tested (if none of the three switches are used, value names will be tested)
        [switch] $ValueName,
        [Parameter(ParameterSetName="SingleSearchString")]
        # When the -SearchRegex parameter is used, this switch means that the value data will be tested (if none of the three switches are used, value data will be tested)
        [switch] $ValueData,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against key names only
        [string] $KeyNameRegex,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against value names only
        [string] $ValueNameRegex,
        [Parameter(ParameterSetName="MultipleSearchStrings")]
        # Specifies a regex that will be checked against value data only
        [string] $ValueDataRegex
    )

    begin {
        switch ($PSCmdlet.ParameterSetName) {
            SingleSearchString {
                $NoSwitchesSpecified = -not ($PSBoundParameters.ContainsKey("KeyName") -or $PSBoundParameters.ContainsKey("ValueName") -or $PSBoundParameters.ContainsKey("ValueData"))
                if ($KeyName -or $NoSwitchesSpecified) { $KeyNameRegex = $SearchRegex }
                if ($ValueName -or $NoSwitchesSpecified) { $ValueNameRegex = $SearchRegex }
                if ($ValueData -or $NoSwitchesSpecified) { $ValueDataRegex = $SearchRegex }
            }
            MultipleSearchStrings {
                # No extra work needed
            }
        }
    }

    process {
        foreach ($CurrentPath in $Path) {
            Get-ChildItem $CurrentPath -Recurse:$Recurse | 
                ForEach-Object {
                    $Key = $_

                    if ($KeyNameRegex) { 
                        Write-Verbose ("{0}: Checking KeyNamesRegex" -f $Key.Name) 
        
                        if ($Key.PSChildName -match $KeyNameRegex) { 
                            Write-Verbose "  -> Match found!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "KeyName"
                            }
                        } 
                    }
        
                    if ($ValueNameRegex) { 
                        Write-Verbose ("{0}: Checking ValueNamesRegex" -f $Key.Name)
            
                        if ($Key.GetValueNames() -match $ValueNameRegex) { 
                            Write-Verbose "  -> Match found!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "ValueName"
                            }
                        } 
                    }
        
                    if ($ValueDataRegex) { 
                        Write-Verbose ("{0}: Checking ValueDataRegex" -f $Key.Name)
            
                        if (($Key.GetValueNames() | % { $Key.GetValue($_) }) -match $ValueDataRegex) { 
                            Write-Verbose "  -> Match!"
                            return [PSCustomObject] @{
                                Key = $Key
                                Reason = "ValueData"
                            }
                        }
                    }
                }
        }
    }
}
## End Search-Registry
## Begin Get-RegistryKeyPropertiesAndValues
Function Get-RegistryKeyPropertiesAndValues{
  <#
    This Function is used here to retrieve registry values while omitting the PS properties
    Example: Get-RegistryKeyPropertiesAndValues -path 'HKCU:\Volatile Environment'
    Origin: Http://www.ScriptingGuys.com/blog
    Via: http://stackoverflow.com/questions/13350577/can-powershell-get-childproperty-get-a-list-of-real-registry-keys-like-reg-query
  #>

 Param(
  [Parameter(Mandatory=$true)]
  [string]$path
  )

  Push-Location
  Set-Location -Path $path
  Get-Item . |
  Select-Object -ExpandProperty property |
  ForEach-Object {
      New-Object psobject -Property @{"Folder"=$_;
        "RedirectedLocation" = (Get-ItemProperty -Path . -Name $_).$_}}
  Pop-Location

# Get the user profile path, while escaping special characters because we are going to use the -match operator on it
$Profilepath = [regex]::Escape($env:USERPROFILE)

# List all folders
$RedirectedFolders = Get-RegistryKeyPropertiesAndValues -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" | Where-Object {$_.RedirectedLocation -notmatch "$Profilepath"}
if ($RedirectedFolders -eq $null) {
    Write-Output "No folders are redirected for this user"
} else {
    $RedirectedFolders | format-list *
}}
## End Get-RegistryKeyPropertiesAndValues
## Begin Get-RegKeyandValues
Function Get-RegKeyandValues{
###########################################################
#Script Title: Get Registry Key and Values PowerShell Tool
#Script File Name: Get-RegKeyandValues.ps1
#Author: Ron Ratzlaff (aka "The_Ratzenator")
#Date Created: 6/22/2014
###########################################################

#Requires -Version 3.0
   <#
	  .SYNOPSIS
	  
	  	The "Get-RegKeyandValues" Function will attempt to retrieve Registry keys and values that you specify, if they exist 
	  
	  .EXAMPLE

		To get the values and the sub keys within the "MyRegKey" key on the local computer, use the following syntax:

		Get-RegKeyandValues -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .EXAMPLE
		
		To get the values, but exclude the sub keys within the "MyRegKey" key on the local computer, use the following syntax:
		
		Get-RegKeyandValues -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "Yes" -GetRegSubKeys "No"

	  .EXAMPLE

		To get the sub keys, but exclude the values within the "MyRegKey" key on the local computer, use the following syntax:
		
		Get-RegKeyandValues -ComputerName "Computer1" -RegHive "HKLM" -RegKey "SOFTWARE\MyRegKey" -GetRegKeyVals "No" -GetRegSubKeys "Yes"
	
	  .EXAMPLE
		
		To retrieve the Registry info remotely on more than one computer, you can use an array as shown in the following syntax:
		
		Get-RegKeyandValues -ComputerName @("Computer1", "Computer2") -RegHive "HKLM" -RegKey "SOFTWARE\Wow6432Node" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .EXAMPLE
		
		To retrieve the Registry info remotely on more than one computer, you can use a file (.csv, .txt) and use the Get-Content cmdlet as shown in the following syntax:
		
		Get-RegKeyandValues -ComputerName (Get-Content -Path "$env:TEMP\ComputerList.txt") -RegHive "HKLM" -RegKey "SOFTWARE\Wow6432Node" -GetRegKeyVals "Yes" -GetRegSubKeys "Yes"

	  .PARAMETER ComputerName
	  
	  	A mandatory parameter used to query a single computer or multiple computers.
	  
	  .PARAMETER RegHive
	  
	  	A mandatory parameter used to query one of the primary Registry Hives (HKCR, HKCU, HKLM, HKU, and HKCC). You must specify one of the primary Registry Hives as shown below, or an error will display:
		
		HKCR: HKEY_CLASSES_ROOT
		HKCU: HKEY_CURRENT_USER
		HKLM: HKEY_LOCAL_MACHINE
		HKU: HKEY_USERS
		HKCC: HKEY_CURRENT_CONFIG
	  
	  .PARAMETER RegKey
	  
	  	A manadatory parameter used to query a specified key. A path must be specified for the key. For instance, to query the "Uninstall" key located under HKLM, you will need to specify the following path:
		
			"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	  
	  .PARAMETER GetRegKeyVals
	  
	  	An optional parameter (not mandatory) that retrieves the values listed within the specified Registry Key parameter (-RegKey) that when used, requires that either a "Yes" or a "No" value is specified
	  
	  .PARAMETER GetRegSubKeys
	  	
		An optional parameter (not mandatory) that retrieves the sub keys listed within the specified Registry Key parameter (-RegKey) that when used, requires that either a "Yes" or a "No" value is specified
  #>
   
   [cmdletbinding()]            
	
	Param
	(            
	 	[Parameter(Position=0,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What computer name would you like to target?')]
	  	$ComputerName = $env:COMPUTERNAME,
		
		[Parameter(Mandatory=$true,
	  	 Position=1,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What Registry Hive would you like to target?')]
		 [ValidateSet('HKCR','HKCU','HKLM','HKU',IgnoreCase = $true)]
		 [ValidateNotNullOrEmpty()]
		[string[]]$RegHive,
		
		[Parameter(Mandatory=$true,
		 Position=2,
		 ValueFromPipeline,
	  	 ValueFromPipelineByPropertyName,
	  	 HelpMessage='What Registry Key would you like to target?')]
		 [ValidateNotNullOrEmpty()] 
		[string[]]$RegKey,
		
		[Parameter(Position=3,
	  	 HelpMessage='Would you like to display the Registry Values within the "$RegKey" Key?')]
		[ValidateSet('Yes', 'No', IgnoreCase = $true)]
		[string[]]$GetRegKeyVals,
		
		[Parameter(Position=4,
	  	 HelpMessage='Would you like to display the Registry Sub Keys under the "$RegKey" Key?')]
		[ValidateSet('Yes', 'No',IgnoreCase = $true)]
		[string[]]$GetRegSubKeys
	)   
   
    Begin {}
   
    Process
	{
		$NewLine = "`r`n"
		
		Switch ($RegHive)
		{
			"HKCR"
			{
				$Hive = "ClassesRoot"
			}
			
			"HKCU"
			{
				$Hive = "CurrentUsers"
			}
			
			"HKLM"
			{
				$Hive = "LocalMachine"
			}
			
			"HKU"
			{
				$Hive = "Users"
			}
			
			"HKCC"
			{
				$Hive = "CurrentConfig"
			}
		}
		
		Foreach ($Computer in $ComputerName)
		{
			$RegHiveType = [Microsoft.Win32.RegistryHive]::$Hive
			$OpenBaseRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHiveType, $Computer)
			$OpenRegSubKey = $OpenBaseRegKey.OpenSubKey($RegKey)
			
			If (Test-Connection -ComputerName $Computer -Count 1 -Quiet)
			{
				If ($OpenRegSubKey)
				{			
					If ($GetRegKeyVals -match "Yes" -and $GetRegSubKeys -match "Yes")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Key Values"
						Write-Output "--------------"
						
						$NewLine
						
						$GetRegKeyVal = Foreach($RegKeyValue in $OpenRegSubKey.GetValueNames()){$RegKeyValue}
						
						If ($GetRegKeyVal -ne $null)
						{
							$GetRegKeyVal
						}
						
						Else
						{
							Write-Output "No Registry Values exist within the $RegKey"
						}
						
						$NewLine
						
						Write-Output "Reg Sub Keys"
						Write-Output "------------"
						
						$NewLine
						
						$GetRegSubKey = Foreach($RegSubKey in $OpenRegSubKey.GetSubKeyNames()){$RegSubKey}
						
						If ($GetRegSubKey -ne $null)
						{
							$GetRegSubKey
						}
						
						Else
						{
							Write-Output "No Registry Sub Keys exist under the $RegKey"
						}
						
						$NewLine
					}
					
					ElseIf ($GetRegKeyVals -match "Yes" -and $GetRegSubKeys -match "No")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Key Values"
						Write-Output "--------------"
						
						$NewLine
						
						$GetRegKeyVal = Foreach($RegKeyValue in $OpenRegSubKey.GetValueNames()){$RegKeyValue}
						
						If ($GetRegKeyVal -ne $null)
						{
							$GetRegKeyVal
						}
						
						Else
						{
							Write-Output "No Registry Values exist within the $RegKey"
						}
						
						$NewLine
					}
					
					ElseIf ($GetRegKeyVals -match "No" -and $GetRegSubKeys -match "Yes")
					{
						$NewLine 
					
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
						
						Write-Output "Reg Sub Keys"
						Write-Output "------------"
						
						$NewLine
						
						$GetRegSubKey = Foreach($RegSubKey in $OpenRegSubKey.GetSubKeyNames()){$RegSubKey}
						
						If ($GetRegSubKey -ne $null)
						{
							$GetRegSubKey
						}
						
						Else
						{
							Write-Output "No Registry Sub Keys exist under the $RegKey"
						}
						
						$NewLine
					}
					
					Else
					{
						$NewLine
						
						Write-Output "Computer"
						Write-Output "--------"
						
						$NewLine
						
						Write-Output "$Computer"
						
						$NewLine
						
						Write-Output "Reg Key Exist"
						Write-Output "-------------"
						
						$NewLine
						
						Write-Output "'$RegKey'"
						
						$NewLine
					}
				}
				
				Else
				{
					$NewLine
					
					Write-Output "Computer"
					Write-Output "--------"
					
					$NewLine
					
					Write-Output "$Computer"
					
					$NewLine
					
					Write-Warning -Message "Could not find $RegKey"
				
					$NewLine
				}
			}
			
			Else
			{
				$NewLine
				
				Write-Warning -Message "$Computer is offline!"
				
				$NewLine
			}
		}
	}
	
	End {}
}#EndFunction
## End Get-RegKeyandValues
## Begin Get-RemoteRegistry
Function Get-RemoteRegistry {
########################################################################################
## Version: 2.1
##  + Fixed a pasting bug 
##  + I added the "Properties" parameter so you can select specific registry values
## NOTE: you have to have access, and the remote registry service has to be running
########################################################################################
## USAGE:
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP"
##     * Returns a list of subkeys (because this key has no properties)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727"
##     * Returns a list of subkeys and all the other "properties" of the key
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\Version"
##     * Returns JUST the full version of the .Net SP2 as a STRING (to preserve prior behavior)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version
##     * Returns a custom object with the property "Version" = "2.0.50727.3053" (your version)
##   Get-RemoteRegistry $RemotePC "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727" Version,SP
##     * Returns a custom object with "Version" and "SP" (Service Pack) properties
##
##  For fun, get all .Net Framework versions (2.0 and greater) 
##  and return version + service pack with this one command line:
##
##    Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP" | 
##    Select -Expand Subkeys | ForEach-Object { 
##      Get-RemoteRegistry $RemotePC "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\$_" Version,SP 
##    }
##
########################################################################################
param(
    [string]$computer = $(Read-Host "Remote Computer Name")
   ,[string]$Path     = $(Read-Host "Remote Registry Path (must start with HKLM,HKCU,etc)")
   ,[string[]]$Properties
   ,[switch]$Verbose
)
if ($Verbose) { $VerbosePreference = 2 } # Only affects this script.

   $root, $last = $Path.Split("\")
   $last = $last[-1]
   $Path = $Path.Substring($root.Length + 1,$Path.Length - ( $last.Length + $root.Length + 2))
   $root = $root.TrimEnd(":")

   #split the path to get a list of subkeys that we will need to access
   # ClassesRoot, CurrentUser, LocalMachine, Users, PerformanceData, CurrentConfig, DynData
   switch($root) {
      "HKCR"  { $root = "ClassesRoot"}
      "HKCU"  { $root = "CurrentUser" }
      "HKLM"  { $root = "LocalMachine" }
      "HKU"   { $root = "Users" }
      "HKPD"  { $root = "PerformanceData"}
      "HKCC"  { $root = "CurrentConfig"}
      "HKDD"  { $root = "DynData"}
      default { return "Path argument is not valid" }
   }


   #Access Remote Registry Key using the static OpenRemoteBaseKey method.
   Write-Verbose "Accessing $root from $computer"
   $rootkey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($root,$computer)
   if(-not $rootkey) { Write-Error "Can't open the remote $root registry hive" }

   Write-Verbose "Opening $Path"
   $key = $rootkey.OpenSubKey( $Path )
   if(-not $key) { Write-Error "Can't open $($root + '\' + $Path) on $computer" }

   $subkey = $key.OpenSubKey( $last )
   
   $output = new-object object

   if($subkey -and $Properties -and $Properties.Count) {
      foreach($property in $Properties) {
         Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
      }
      Write-Output $output
   } elseif($subkey) {
      Add-Member -InputObject $output -Type NoteProperty -Name "Subkeys" -Value @($subkey.GetSubKeyNames())
      foreach($property in $subkey.GetValueNames()) {
         Add-Member -InputObject $output -Type NoteProperty -Name $property -Value $subkey.GetValue($property)
      }
      Write-Output $output
   }
   else
   {
      $key.GetValue($last)
   }
 }
## Begin Get-RemoteRegistryKey
Function Get-RemoteRegistryKey {
    <#
    .SYNOPSIS
        Set registry key on remote computers.
    .DESCRIPTION
        This Function uses .Net class [Microsoft.Win32.RegistryKey].
    .PARAMETER ComputerName
        Name of the remote computers.
    .PARAMETER Hive
        Hive where the key is.
    .PARAMETER KeyPath
        Path of the key.
    .PARAMETER Name
        Name of the key setting.
    .PARAMETER Type
        Type of the key setting.
    .PARAMETER Value
        Value tu put in the key setting.
    .EXAMPLE
        Get-RemoteRegistryKey -ComputerName $env:ComputerName -Hive "LocalMachine" -KeyPath "software\corporate\master\Test" -Name "TestName" -Type String -Value "TestValue" -Verbose
    .LINK
        http://itfordummies.net
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage='Provide a ComputerName')]
        [String[]]$ComputerName=$env:ComputerName,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryHive]$Hive,
        
        [Parameter(Mandatory=$true)]
        [String]$KeyPath,
        
        [Parameter(Mandatory=$true)]
        [String]$Name,
        
        [Parameter(Mandatory=$true)]
        [Microsoft.Win32.RegistryValueKind]$Type,
        
        [Parameter(Mandatory=$true)]
        [Object]$Value
    )
    Begin{
    }
    Process{
        ForEach ($Computer in $ComputerName) {
            try {
                Write-Verbose "Trying computer $Computer"
                $reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("$hive", "$Computer")
                Write-Debug -Message "Contenur de Reg $reg"
                $key=$reg.OpenSubKey("$KeyPath",$true)
                if($key -eq $null){
                    Write-Verbose -Message "Key not found."
                    Write-Verbose -Message "Calculating parent and child paths..."
                    $parent = Split-Path -Path $KeyPath -Parent
                    $child = Split-Path -Path $KeyPath -Leaf
                    Write-Verbose -Message "Creating the subkey $child in $parent..."
                    $Key=$reg.OpenSubKey("$parent",$true)
                    $Key.CreateSubKey("$child") | Out-Null
                    Write-Verbose -Message "Setting $value in $KeyPath"
                    $key=$reg.OpenSubKey("$KeyPath",$true)
                    $key.SetValue($Name,$Value,$Type)
                }
                else{
                    Write-Verbose "Key found, setting $Value in $KeyPath..."
                    $key.SetValue($Name,$Value,$Type)
                }
                Write-Verbose "$Computer done."
            }#End Try
            catch {Write-Warning "$Computer : $_"} 
        }#End ForEach
    }#End Process
    End{
    }
}
## End Get-RemoteRegistryKey