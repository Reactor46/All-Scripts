###########################################################
#Script Title: Get Registry Key and Values PowerShell Tool
#Script File Name: Get-RegKeyandValues.ps1
#Author: Ron Ratzlaff (aka "The_Ratzenator")
#Date Created: 6/22/2014
###########################################################

#Requires -Version 3.0

Function Get-RegKeyandValues
{
   <#
	  .SYNOPSIS
	  
	  	The "Get-RegKeyandValues" function will attempt to retrieve Registry keys and values that you specify, if they exist 
	  
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
	