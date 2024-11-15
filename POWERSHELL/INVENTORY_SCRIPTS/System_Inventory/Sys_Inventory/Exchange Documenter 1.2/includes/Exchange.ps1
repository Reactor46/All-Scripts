#############################################################################
#                                     			 		    #
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 		    #
#############################################################################



#region view settings
# We want to view the Entire Forest
Import-Module ActiveDirectory
try{
$ErrorActionPreference = "Stop"
$SRVSettings = Get-ADServerSettings
if ($SRVSettings.ViewEntireForest -eq "False")
	{   #Add-text "Changing settings to view Entire Forest" #-ForegroundColor Green
		Set-ADServerSettings -ViewEntireForest $true
	}
}
Catch {
 #Add-text $_.Message #-ForegroundColor DarkRed
}
#endregion view settings

#region Organization Name
Function Get-Org 
{
try{
    $ErrorActionPreference = "Stop"
    Get-OrganizationConfig | select Name | Export-Csv ".\data\orgname.csv"
    }
    Catch {
    #Add-text $_.Message # -ForegroundColor Red
    }
}
#endregion Organization Name

#region Exchange Servers Name
Function Get-Exservers
{
$Results = @()
$Items =  @(Get-ExchangeServer | where {($_.ServerRole -ne "Edge" -and $ExcludedServers -notcontains $_.Name) -or ($_.IsEdgeServer -ne "True" -and $ExcludedServers -notcontains $_.Name) } | Select Identity,AdminDisplayVersion,Site,Edition -Unique)
	
	Foreach ($i in $Items)
	{
		IF (@($ListCheck.CheckedItems) -notcontains  $I.Identity)
		{
		$obj = New-Object PSObject
		$obj | Add-Member NoteProperty -Name "Identity" -Value $I.Identity
		$obj | Add-Member NoteProperty -Name "AdminDisplayVersion" -Value $I.AdminDisplayVersion
		$obj | Add-Member NoteProperty -Name "Site" -Value $I.Site.Name
		$obj | Add-Member NoteProperty -Name "Edition" -Value $I.Edition
		$results += $obj
		}
	}
	
	Return $Results
}

#endregion Exchange Servers Name

#region Exchange Servers Roles
Function get-ExRoles
{
#Add-text "Getting Exchange Servers Roles" #-ForegroundColor Green
   Get-ExchangeServer  |  where {($_.ServerRole -ne "Edge" -and $ExcludedServers -notcontains $_.Name) -or ($_.IsEdgeServer -ne "True" -and $ExcludedServers -notcontains $_.Name) }  | Select Identity,ServerRole,AdminDisplayVersion
}
#endregion Exchange Servers Roles

#region Exchange Journal
Function get-ExJournal
{
#Try{
$ErrorActionPreference = "Stop"
$return = @()
    #Add-text "Collecting Journal recipient information" 
    $DBJournal = @(Get-MailboxDatabase -inc |?{$ExcludedServers -notcontains $_.ServerName  })
    Foreach ($DB in $DbJournal)
    {
    Try {
     IF ($DB.JournalRecipient -ne $null)
     {
     $Mailrcpt = (Get-Mailbox $DB.JournalRecipient).Name
	 $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "ServerName" -Value $DB.ServerName
     $obj | Add-Member NoteProperty -Name "Name" -Value $DB.Name
     $obj | Add-Member NoteProperty -Name "JournalRecipient" -Value $Mailrcpt
	 $return += $obj  
     }
     
     }
     Catch {
     #Add-text $_.Message #-ForegroundColor Red
     }
	 
    }
		 
	 IF (($return.count) -eq 0)
	 {
	 $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "JournalRecipient" -Value "No Journal Recipients found"
	 $return += $obj  
	 }
return $return
#}
#Catch {$_}
}
#endregion Exchange Journal

#region Exchange Quota

Function get-Quota
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting Quota limits" #-ForegroundColor Green
    $Items = @( Get-Mailboxdatabase -inc  |?{ $ExcludedServers -notcontains $_.ServerName  })
   
   Foreach ($Item in $Items)
   {
   Try {
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name
     $obj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota" -Value $Item.ProhibitSendReceiveQuota
     $obj | Add-Member NoteProperty -Name "ProhibitSendQuota" -Value $Item.ProhibitSendQuota
     $obj | Add-Member NoteProperty -Name "IssueWarningQuota" -Value $Item.IssueWarningQuota
     $Result += $obj
      }
   Catch{ #Add-text $_.Message #-ForegroundColor DarkYellow
   }
   }
 
   Return $result
}

#endregion Exchange Quota

#region Exchange Retention
Function get-Retention
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting Retention settings" # -ForegroundColor Green
    $Items = @(Get-Mailboxdatabase -inc |?{$ExcludedServers -notcontains $_.ServerName })
   
   Foreach ($Item in $Items)
   {
   Try {
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name
     $obj | Add-Member NoteProperty -Name "MailboxRetention" -Value $Item.MailboxRetention
     $obj | Add-Member NoteProperty -Name "DeletedItemRetention" -Value $Item.DeletedItemRetention
     $Result += $obj
      }
   Catch{ #Add-text $_.Message #-ForegroundColor DarkYellow
   }
   }
   Return $result
}
#endregion Exchange Retention

#region Exchange MailboxDB
Function Get-MailboxDB
{
#Add-text "Getting Mailbox Databases" #-ForegroundColor Green
$ErrorActionPreference = "Stop"
$result = @()
Try {
$Databases = @(Get-MailboxDatabase -inc |?{$ExcludedServers -notcontains $_.ServerName } | select Name,EdbFilePath,LogFolderPath,CircularLoggingEnabled)
}
Catch{#Add-text $_.Message #-ForegroundColor DarkRed
}
try{
    Foreach ($MBD in $Databases)
    {
    
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "Name" -Value $MBD.Name
     $obj | Add-Member NoteProperty -Name "EdbFilePath" -Value $MBD.EdbFilePath
     $obj | Add-Member NoteProperty -Name "LogFolderPath" -Value $MBD.LogFolderPath
     $obj | Add-Member NoteProperty -Name "CircularLoggingEnabled" -Value $MBD.CircularLoggingEnabled
     $Result += $obj
       }
     Return $Result 
    }   
       Catch{
      # Add-text $_.Message #-ForegroundColor #DarkRed
       }
}

#endregion Exchange MailboxDB

#region Exchange MailboxCount
Function Get-MailboxCount
{
#Add-text "Calculating mailbox counts" # -ForegroundColor Green
$ErrorActionPreference = "Stop"
$result = @()
Try{
$Databases = @(Get-MailboxDatabase -inc|?{$ExcludedServers -notcontains $_.ServerName }| select Name,ServerName,Identity )
}
Catch{#Add-text $_.Message #-ForegroundColor DarkRed
}

try{
    Foreach ($MBD in $Databases)
    {
     [string]$identity = $MBD.Identity.tostring()
     $count = @(Get-Mailbox -database $identity  -ResultSize Unlimited).count
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "ServerName" -Value $MBD.ServerName
     $obj | Add-Member NoteProperty -Name "Name" -Value $MBD.Name
     $obj | Add-Member NoteProperty -Name "MailboxCount" -Value $count
     $Result += $obj
       }
   Return $Result 
   }
   Catch {#Add-text $_.Message #-ForegroundColor DarkRed
   }
}

#endregion Exchange MailboxCount

#region Exchange rollups
#===================================================================
# Exchange Rollup (Edge server excluded)
#===================================================================
Function Get-ExSP ($DisplayVersion)
{
$version =$DisplayVersion.split(".")[1]
switch ($version) 
    { 
        1 {"Service Pack 1"} 
        2 {"Service Pack 2"} 
        3 {"Service Pack 3"} 
        4 {"Service Pack 4"} 
        5 {"Service Pack 5"} 
        default {"RTM"}
    }
}

Function get-rollups
{
#Add-text "Getting rollup information"
$ErrorActionPreference = "Stop"
try{
$MsxServers = Get-EXServerNodes 
$ClassHeaderRollup = "heading1"
#Loop through each Exchange server that is found

$DetailRollup =  @()
if ($MsxServers -ne $NULL)
{
    ForEach ($MsxServer in $MsxServers)
	{
	$ping = New-Object –TypeName System.Net.Networkinformation.Ping
    Try { 
        $status = ($ping.Send($($MsxServer.Name))).Status 
#Write-Host "$status" -ForegroundColor Yellow
        #Status Check returned with true
        }
    Catch 
	    {
        $status = "Failure"
        #Add-text "Cannot reach server $($MsxServer.Name) . It will be excluded from Rollup and Patch results." #-ForegroundColor red
        }
			If ( $? -eq $true -and  $status -eq "Success")
	        {		
		   	    #Get Exchange server version
		        $vkey = "15"
			    $MsxVersion = $MsxServer.ExchangeVersion
		        IF ($MsxServer.AdminDisplayVersion.major -eq 8)
		        {
		        $vkey = ""
		        $key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\461C2B4266EDEF444B864AD6D9E5B613\Patches\"
		        }
		        ElseIF ($MsxServer.AdminDisplayVersion.major -eq 14)
		        {
		        $vkey = " v14"
		        $key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\AE1D439464EB1b8488741FFA028E291C\Patches\"
		        }
		        ElseIF ($MsxServer.AdminDisplayVersion -like "Version 15.*")
		        {
		        $vkey = " v15"
		        $key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Exchange v15\"
		        }

			    #Create "header" string for output
				
				Try {
#Write-Host "$vkey Version 15 $($MsxServer.Name)" -ForegroundColor Yellow
			    $Srv = $MsxServer.Name
		        $type = [Microsoft.Win32.RegistryHive]::LocalMachine
				$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
			    $regKey = $regKey.OpenSubKey($key)
		            if ($vkey -eq " v15") 
		                    { 
#Write-Host "Version 15" -ForegroundColor Yellow
		                    $obj = New-Object PSObject
		                    $obj | Add-Member NoteProperty -Name "ServerName" -Value $Srv -Force
		                    $obj | Add-Member NoteProperty -Name "DisplayName" -Value $regKey.GetValue("DisplayName") -Force
		                    $DetailRollup+= $Obj
		                    }
		                    Else
		                    {
		                     if ($regkey -ne $null)
			                    {  
		                            if ($regKey.SubKeyCount -eq 0)
		                            {
		                            $key = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Exchange$($vkey)\"
		                            $type = [Microsoft.Win32.RegistryHive]::LocalMachine
			                        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
		                            $regKey = $regKey.OpenSubKey($key)
		                                $obj = New-Object PSObject
		                                $obj | Add-Member NoteProperty -Name "ServerName" -Value $Srv -Force
		                                $obj | Add-Member NoteProperty -Name "DisplayName" -Value "$($regKey.GetValue('DisplayName')) $(Get-ExSP $regKey.GetValue('DisplayVersion'))" -Force
		                            $DetailRollup+= $obj
#Write-Host "Object $($obj.ToString())" -ForegroundColor Yellow
		                            #$DetailRollup+=  "NO ROLLUP INSTALLED"
		                            }
		                            Else
		                            {
		                            #Loop each of the subkeys (Patches) and gather the Installed date and Displayname of the Exchange 2007 patch
			                        $ErrorActionPreference = "SilentlyContinue"
			                        ForEach($sub in $regKey.GetSubKeyNames())
			                                    {
				                                    $SUBkey = $key + $Sub
				                                    $SUBregKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $Srv)
				                                    $SUBregKey = $SUBregKey.OpenSubKey($SUBkey)

				                                              $obj = New-Object PSObject
		                                                      $obj | Add-Member NoteProperty -Name "ServerName" -Value $Srv -Force
		                                                      $obj | Add-Member NoteProperty -Name "DisplayName" -Value $SUBRegkey.GetValue('DisplayName') -Force
		                                                      $DetailRollup+= $obj
		                                                     
			                                    }
		                             }
		                         }
		                    }
				} Catch [Exception] {
#Add-text "Cannot retrieve Patch Information for $($Srv) "
}
			}
    }
	


}
} #end try
Catch {
$_
}
#Return $DetailRollup
Format-Rollup $DetailRollup
}
#endregion Exchange rollups

#region Exchange CasArray
Function get-CASArray
{
$ErrorActionPreference = "Stop"
$Report = @()
    #Add-text "Getting CAS array information"
	If ($ShellVersion -eq "8")
	{
	 #Add-text "ClientAccessArray does not exist in Exchange 2007"
	 $obj = New-Object PSObject
	 $obj | Add-Member NoteProperty -Name "Members" -Value "N/A"
     $obj | Add-Member NoteProperty -Name "SiteName" -Value "N/A"
     $obj | Add-Member NoteProperty -Name "Identity" -Value "N/A"
	 $Report += $obj
	}
	Else
	{
	Try{
    $Items = @(Get-ClientAccessarray)
   }
   Catch{
   #Add-text $_.message
   }
   Foreach ($Item in $Items)
   {
		   Try {
		     $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "FQDN" -Value $Item.FQDN
		     $obj | Add-Member NoteProperty -Name "Members" -Value ($item.Members -join " ")
		     $obj | Add-Member NoteProperty -Name "SiteName" -Value $Item.SiteName
		     $obj | Add-Member NoteProperty -Name "Identity" -Value $Item.Identity
		     $Report += $obj
		      }
   Catch{ 
   #Add-text $_.Message
   }
   }
   }
	Return $Report
}
#endregion Exchange CasArray

#region Exchange CasSettings
Function CAS-Settings
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting Client Access Server information"# -ForegroundColor Green
    $Items = @(Get-ClientAccessServer |?{$ExcludedServers -notcontains $_.Name })
   
   Foreach ($Item in $Items)
   {
   IF (@($ListCheck.CheckedItems) -notcontains  $Item.Name)
   {
   Try {
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name
     $obj | Add-Member NoteProperty -Name "OutlookAnywhereEnabled" -Value $Item.OutlookAnywhereEnabled
     $obj | Add-Member NoteProperty -Name "AutoDiscoverServiceInternalUri" -Value $Item.AutoDiscoverServiceInternalUri
     $obj | Add-Member NoteProperty -Name "AutoDiscoverSiteScope" -Value ($Item.AutoDiscoverSiteScope -join " ")
     $Result += $obj
      }
   Catch{ #Add-text $_.Message -ForegroundColor #DarkYellow
   }
   }
   }
  Return $result
  
}
#endregion Exchange CasSettings

#region Exchange OWA
Function Get-OWAVDIR
{
$ErrorActionPreference = "Stop"
$result = @()
    #Write-Host "Getting OWA information" -ForegroundColor Green
	try{
    $Items = @(Get-ExchangeServer | ?{($_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name)} | select -Unique) 
      }
   Catch{$_}
   Foreach ($Item in $Items)
   {
   IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
   {
   Try {
   $Server = @(Get-OwaVirtualDirectory -Server $item.Identity | ?{$_.Displayname -eq 'OWA'}  )
   Foreach ($site in $server)
     {
     $obj = New-Object PSObject
    
     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value $site.BasicAuthentication
     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value $site.DigestAuthentication
     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value $site.FormsAuthentication
     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value $site.WindowsAuthentication
     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
     $Result += $obj
     }
      }
   Catch{ 
   #Write-host $_.Message -ForegroundColor DarkYellow
   }
   }
}
   Return $result
 
}

#endregion Exchange OWA

#region Exchange OAB
Function Get-OABVDIR
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting OAB information" #-ForegroundColor Green
	try{
    $Items = @(Get-ExchangeServer | ?{($_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name) | select -Unique})
      }
   Catch{$_}
   Foreach ($Item in $Items)
   {
      IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
	  {
	   Try {
	   $Server = @(Get-OabVirtualDirectory -server $Item.Identity)
   Foreach ($site in $server)
     {
	     $obj = New-Object PSObject
	     #$obj | Add-Member NoteProperty -Name "Server" -Value $Item.Server
	     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value $site.BasicAuthentication
	     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value $site.DigestAuthentication
	     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value $site.FormsAuthentication
	     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value $site.WindowsAuthentication
	     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
	     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
	     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
	     $Result += $obj
        }
	      }
	   Catch{ #Write-host $_.Message -ForegroundColor DarkYellow
	   }
	   }
   }

   Return $result
  
}
#endregion Exchange OAB

#region Exchange EWS
Function Get-EWS
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting EWS information" #-ForegroundColor Green
	try{
    $Items = @(Get-ExchangeServer | ?{($_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name) | Select -Unique})
   }
   catch{$_}
   
   Foreach ($Item in $Items)
   {
	      IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
		  {
	   Try {
	   $Server = @(Get-WebServicesVirtualDirectory  -server $Item.Identity)
   Foreach ($site in $server)
     {
	     $obj = New-Object PSObject
	     #$obj | Add-Member NoteProperty -Name "Server" -Value $Item.Server
	     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value $site.BasicAuthentication
	     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value $site.DigestAuthentication
	     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value $site.FormsAuthentication
	     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value $site.WindowsAuthentication
	     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
	     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
	     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
	     $Result += $obj
}
	      }
	   Catch{ #Add-text $_.Message #-ForegroundColor DarkYellow
	   }
	   }
	}
   Return $result
  
}

#endregion Exchange EWS

#region ExchangeActiveSync

Function Get-ASVDIR
{
$ErrorActionPreference = "Stop"
$result = @()
    #Add-text "Getting ActiveSync information" #-ForegroundColor Green
	try{
    $Items = @(Get-ExchangeServer | ?{($_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name) | select -Unique})
   }
   Catch{$_}
   Foreach ($Item in $Items)
   {
      IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
	  {
   Try {
   $Server = @(Get-ActiveSyncVirtualDirectory -Server $Item.Identity)
      Foreach ($site in $server)
     {
     $obj = New-Object PSObject
     $obj | Add-Member NoteProperty -Name "Server" -Value $site.Server
     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
     $Result += $obj
     }
      }
   Catch {#Add-text $_.Message #-ForegroundColor DarkYellow
   }
   }
 }
   Return $result
   
   
}

#endregion ExchangeActiveSync

#region ExchangeAutoDiscover

Function Get-AutoDS
{
$ErrorActionPreference = "Stop"
$Report = @()
    #Add-text "Getting AutoDiscover information" #-ForegroundColor Green
	try{
    $Items = @(Get-ExchangeServer | ?{ ($_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name) | select -Unique})
    }
   Catch{}  
   Foreach ($Item in $Items)
   {
      IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
	  {
   Try {
    $Server = @(Get-AutodiscoverVirtualDirectory -server $Item.Identity)
          Foreach ($site in $server)
     {
     $obj = New-Object PSObject
     #$obj | Add-Member NoteProperty -Name "Server" -Value $Server.Server
     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value $site.BasicAuthentication
     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value $site.DigestAuthentication
     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value $site.FormsAuthentication
     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value $site.WindowsAuthentication
     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
     $Report += $obj
     }
      }
   Catch {#Add-text $_.Message # -ForegroundColor DarkYellow
   }
   }
   }
   Return $Report
}
#endregion ExchangeAutoDiscover

#region Exchange ECP
Function Get-ECP
{
$ErrorActionPreference = "Stop"
$Report = @()
    #Write-Host "Getting ECP information" -ForegroundColor Green
	If ($ShellVersion -eq "8")
	{
	#Add-text "cmdlet not vailable in Exchange Server 2007 EMS"
			 $obj = New-Object PSObject
		    # $obj | Add-Member NoteProperty -Name "Server" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "Externalurl" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "Internalurl" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "Identity" -Value "N/A"
		     $Report += $obj
	}
	Else
	{

	TRY {
		    $Items = @(Get-ExchangeServer | ?{($_.AdmindisplayVersion.major -gt 8 -and $_.serverrole -like '*ClientAccess*' -and $ExcludedServers -notcontains $_.Name ) -or ($_.isClientAccessServer -eq 'true' -and $ExcludedServers -notcontains $_.Name) | Select -Unique }) 
		   }
		Catch {
		#$_
		}
		   Foreach ($Item in $Items)
		   {
		   	   IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
	   		{
			   Try {
			     $Server = @(Get-EcpVirtualDirectory -Server $Item.Identity)
          Foreach ($site in $server)
     {
			     $obj = New-Object PSObject
			     $obj | Add-Member NoteProperty -Name "Server" -Value $site.Identity
			     $obj | Add-Member NoteProperty -Name "BasicAuthentication" -Value $site.BasicAuthentication
			     $obj | Add-Member NoteProperty -Name "DigestAuthentication" -Value $site.DigestAuthentication
			     $obj | Add-Member NoteProperty -Name "FormsAuthentication" -Value $site.FormsAuthentication
			     $obj | Add-Member NoteProperty -Name "WindowsAuthentication" -Value $site.WindowsAuthentication
			     $obj | Add-Member NoteProperty -Name "Externalurl" -Value $site.Externalurl
			     $obj | Add-Member NoteProperty -Name "Internalurl" -Value $site.Internalurl
			     $obj | Add-Member NoteProperty -Name "Identity" -Value $site.Identity
			     $Report += $obj
}
			      }
	   			Catch {			}
   			}
			}
	}
		Return $Report
	}
#endregion Exchange ECP

#region Exchange RPC
Function Get-OA
{
$ErrorActionPreference = "Stop"
$Report = @()
try{
$Items = @(Get-OutlookAnywhere |?{ $ExcludedServers -notcontains $_.Server } | select Identity,ExternalClientAuthenticationMethod,InternalClientAuthenticationMethod,IISAuthenticationMethods,ExternalHostname,InternalHostname,ExchangeVersion,Server)
   }
   Catch{
   #$_
   }
Foreach ($item in $Items)
    {
	   IF (@($ListCheck.CheckedItems) -notcontains  $Item.Identity)
	   {
    $obj = New-Object PSObject
    $obj | Add-Member NoteProperty -Name "Identity"    -Value $item.Identity -Force
    IF (@($item.ExternalClientAuthenticationMethod) -ne $null)
	{
	$obj | Add-Member NoteProperty -Name "ExternalClientAuthenticationMethod"  -Value ($item.ExternalClientAuthenticationMethod -join " ") -Force
    }
	IF (@($item.InternalClientAuthenticationMethod) -ne $null)
	{
    $obj | Add-Member NoteProperty -Name "InternalClientAuthenticationMethod"  -Value ($item.InternalClientAuthenticationMethod -join " ") -Force
    }
 	IF (@($item.IISAuthenticationMethods) -ne $null)
	{
    $obj | Add-Member NoteProperty -Name "IISAuthenticationMethods" -Value ($item.IISAuthenticationMethods -join " ") -Force
    }
    $obj | Add-Member NoteProperty -Name "ExternalHostname"   -Value $item.ExternalHostname -Force
    $obj | Add-Member NoteProperty -Name "InternalHostname"         -Value $item.InternalHostname -Force
    $obj | Add-Member NoteProperty -Name "ExchangeVersion"   -Value $item.ExchangeVersion -Force
    $obj | Add-Member NoteProperty -Name "Server"         -Value $item.Server -Force
    $Report += $obj
    }
	}
    Return $Report
}

#endregion Exchange RPC

#region Exchange DAG
Function Get-DAG
{
$ErrorActionPreference = "Stop"
$report = @()
#Add-text "DAG information" 
	If ($ShellVersion -eq "8")
	{
	#Add-text "Database Availability Groups does not exist in exchange 2007"
	$obj = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "StoppedMailboxServers"  -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "WitnessServer" -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "WitnessDirectory"   -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "AlternateWitnessDirectory"         -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabaseCopiesPerVolume"   -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabaseCopiesPerDatabase"         -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabasesRootFolderPath"         -Value "N/A" -Force
    $obj | Add-Member NoteProperty -Name "AutoDagVolumesRootFolderPath"         -Value "N/A" -Force
    $report += $obj
	}
	Else
	{
try{
$Items =  @(Get-DatabaseAvailabilityGroup |Select Name,Servers,WitnessServer,WitnessDirectory,AlternateWitnessDirectory,StoppedMailboxServers,AutoDagDatabaseCopiesPerVolume,AutoDagDatabaseCopiesPerDatabase,AutoDagDatabasesRootFolderPath,AutoDagVolumesRootFolderPath)
   }
   Catch{
   #Add-text $_.Message  
   }
Foreach ($item in $Items)
    {
    $obj = New-Object PSObject
    $obj | Add-Member NoteProperty -Name "Name" -Value $item.Name -Force
	If (@($item.Servers)-ne $null)
	{
    $obj | Add-Member NoteProperty -Name "Servers"  -Value ($item.Servers -join " ") -Force
    }
    If (@($item.StoppedMailboxServers)-ne $null)
	{
	$obj | Add-Member NoteProperty -Name "StoppedMailboxServers"  -Value ($item.StoppedMailboxServers -join " ") -Force
	}
    $obj | Add-Member NoteProperty -Name "WitnessServer" -Value $item.WitnessServer -Force
    $obj | Add-Member NoteProperty -Name "WitnessDirectory"   -Value $item.WitnessDirectory -Force
    $obj | Add-Member NoteProperty -Name "AlternateWitnessDirectory"         -Value $item.AlternateWitnessDirectory -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabaseCopiesPerVolume"   -Value $item.AutoDagDatabaseCopiesPerVolume -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabaseCopiesPerDatabase"         -Value $item.AutoDagDatabaseCopiesPerDatabase -Force
    $obj | Add-Member NoteProperty -Name "AutoDagDatabasesRootFolderPath"         -Value $item.AutoDagDatabasesRootFolderPath -Force
    $obj | Add-Member NoteProperty -Name "AutoDagVolumesRootFolderPath"         -Value $item.AutoDagVolumesRootFolderPath -Force

    $report += $obj
    }
    }
	Return $report
}
#endregion Exchange DAG

#region Exchange Nodes
Function Get-EXServerNodes
{
#Add-text "Server Nodes"
$ErrorActionPreference = "Stop"
try{
$AllNodes = @()
$ExServers = @(Get-ExchangeServer | where {($_.ServerRole -ne "Edge" -and $ExcludedServers -notcontains $_.Name) -or ($_.IsEdgeServer -ne "True" -and $ExcludedServers -notcontains $_.Name) | select -Unique })
ForEach ($srv in $ExServers)
    {
	#IF (@($ListCheck.CheckedItems) -notcontains  $srv.Identity)
	#{
    If ($srv.AdminDisplayVersion.major -eq 8)
            {
            #Checking for Clusters 
            if ($srv.IsMemberOfCluster -eq "yes")
                {
                    $Nodes = (Get-Mailboxserver $Srv.Name).RedundantMachines
                        ForEach ($N in $Nodes)
                        {
                         #Getting Cluster Nodes
                         $OBJ = New-Object PSObject
                         $obj | Add-Member NoteProperty -Name "Name" -Value $N -Force
                         $obj | Add-Member NoteProperty -Name "AdminDisplayVersion" -Value $srv.AdminDisplayVersion -Force
                         $obj | Add-Member NoteProperty -Name "ExchangeVersion" -Value $srv.ExchangeVersion -Force
                         $AllNodes += $obj
                        }
                }
                 Else
                {
                #If It is is not a cluster
                         $OBJ = New-Object PSObject
                         $obj | Add-Member NoteProperty -Name "Name" -Value $srv.Name -Force
                         $obj | Add-Member NoteProperty -Name "AdminDisplayVersion" -Value $srv.AdminDisplayVersion -Force
                         $obj | Add-Member NoteProperty -Name "ExchangeVersion" -Value $srv.ExchangeVersion -Force
                         $AllNodes += $obj
                }
            }
    Else
    {        $OBJ = New-Object PSObject
             $obj | Add-Member NoteProperty -Name "Name" -Value $srv.Name -Force
             $obj | Add-Member NoteProperty -Name "AdminDisplayVersion" -Value $srv.AdminDisplayVersion -Force
             $obj | Add-Member NoteProperty -Name "ExchangeVersion" -Value $srv.ExchangeVersion -Force
             $AllNodes += $obj
    }
     
	}
 return $AllNodes
}
Catch{#Add-text $_.Message #-ForegroundColor DarkRed
}
}

#endregion Exchange Nodes

#region Exchange Schema
Function Get-Schema 
{
#Add-text "Schema"
$ErrorActionPreference = "Stop"
try {
$ExSchemaversion = ([ADSI]("LDAP://CN=ms-Exch-Schema-Version-Pt," + ([ADSI]"LDAP://RootDSE").schemaNamingContext)).rangeUpper.tostring()
#Write-Host "Exchange Schema Version is $($ExSchemaversion)" -ForegroundColor DarkYellow
$OBJ = New-Object PSObject
$obj | Add-Member NoteProperty -Name "Exchange Schema Version" -Value $ExSchemaversion -Force
Return $obj
}
catch{
$ExSchemaversion = "Not Available"
$OBJ = New-Object PSObject
$obj | Add-Member NoteProperty -Name "Exchange Schema Version" -Value ([string]$ExSchemaversion) -Force
#Write-Host "Could not Deteremine Exchange Schema Version" -ForegroundColor DarkRed}
Return $obj
}
}
#endregion Exchange Schema

#region Custom recieve connector
Function Get-RecCon
{
$ErrorActionPreference = "Stop"
    #Add-text "Custom Receive Connectors" 
$Report = @()	
	$items = @(Get-receiveconnector |?{ $ExcludedServers -notcontains $_.Server} |Select-Object -Property "identity","RemoteIPRanges") #| where {$_.identity –notlike “*default*” –and $_.identity –notlike “*client*”} 
	Foreach ($item in $items) 
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "Identity" -Value $Item.identity -Force
	#[string] $Values = ""
	IF (@($Item.RemoteIPRanges) -ne $null)
	{
	$obj | Add-Member NoteProperty -Name "RemoteIPRanges" -Value ($Item.RemoteIPRanges -join " ") -Force
	}
	$Report += $OBJ
	}
	
	Return $Report
}
#endregion

#region Sendconnector
Function Get-SendCon
{
$ErrorActionPreference = "Stop"
    #Add-text "Getting Send Connectors" 
$Report = @()	
	$items = @(Get-SendConnector | Select-Object -Property "identity","AddressSpaces","IsSMTPConnector","MaxMessageSize","Port","Smarthosts","SourceTransportServers","SmartHostAuthMechanism")
	Foreach ($item in $items) 
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "Identity" -Value $Item.identity -Force
	[string] $Values = ""
	$obj | Add-Member NoteProperty -Name "AddressSpaces" -Value ($Item.AddressSpaces -join " ") -Force
	$obj | Add-Member NoteProperty -Name "IsSmtpConnector" -Value $Item.IsSmtpConnector -Force
	$obj | Add-Member NoteProperty -Name "MaxMessageSize" -Value  $Item.MaxMessageSize  -Force
	$obj | Add-Member NoteProperty -Name "Port" -Value  $Item.Port -Force
	$obj | Add-Member NoteProperty -Name "SmartHostAuthMechanism" -Value  $Item.SmartHostAuthMechanism -Force
		IF (@($Item.SmartHosts) -ne $null)
		{$obj | Add-Member NoteProperty -Name "SmartHosts" -Value ($Item.SmartHosts -join " ") -Force}
		ELSE
		{$obj | Add-Member NoteProperty -Name "SmartHosts" -Value " " -Force}
		IF (@($Item.SourceTransportServers) -ne $null)
		{$obj | Add-Member NoteProperty -Name "SourceTransportServers" -Value ($Item.SourceTransportServers -join " ") -Force	}
		ELSE
		{$obj | Add-Member NoteProperty -Name "SourceTransportServers" -Value " " -Force}
		
	$Report += $OBJ
	}
	
	Return $Report
}
#endregion SendConnector

#region Transport Rules

Function Get-TransRule
{
$ErrorActionPreference = "Stop"
    #Add-text "Transport Rules"
$Report = @()	
	$items = @(Get-TransportRule | Select-Object "Identity","State","Priority","Comments","Description")
	Foreach ($item in $items) 
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "Identity" -Value $Item.Identity -Force
	$obj | Add-Member NoteProperty -Name "State" -Value $Item.State -Force
	$obj | Add-Member NoteProperty -Name "Priority" -Value $Item.Priority -Force
	$obj | Add-Member NoteProperty -Name "Comments" -Value $Item.Comments -Force
	$obj | Add-Member NoteProperty -Name "Description" -Value $Item.Description -Force
	$Report += $OBJ
	}
	
	Return $Report
}
#endregion Transport Rules

#region Addresslists

Function getAddresslists 
{
#Add-text "Address Lists"
$Report = @()
$Alist = @(Get-AddressList)
	Foreach ($item in $Alist)
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "DisplayName" -Value $Item.DisplayName -Force
	$obj | Add-Member NoteProperty -Name "Path" -Value $Item.Path -Force
	$obj | Add-Member NoteProperty -Name "RecipientFilter" -Value $Item.RecipientFilter -Force
	$Report += $OBJ
	}
	Return $Report
}

#endregion Addresslists

#region Accepted Domain
Function GetAcceptedDomains
{
#Add-text "Accepted Domains"
$Report = @()
$Accepted = @(Get-AcceptedDomain)
	Foreach ($item in $Accepted)
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "DomainName" -Value $Item.DomainName -Force
	$obj | Add-Member NoteProperty -Name "DomainType" -Value $Item.DomainType -Force
	$obj | Add-Member NoteProperty -Name "Default" -Value $Item.Default -Force
	$obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name -Force
	$Report += $OBJ
	}
	Return $Report
}
#endregion Accepted Domain

#region OfflineAddressbook
Function OfflineAddressbook
{
#Add-text "OAB"
$Report = @()
$OABS = @(Get-OfflineAddressBook)
	Foreach ($item in $OABS)
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "Server" -Value $Item.Server -Force
	$obj | Add-Member NoteProperty -Name "AddressLists" -Value ($Item.AddressLists  -join " ") -Force
	$obj | Add-Member NoteProperty -Name "IsDefault" -Value $Item.IsDefault -Force
	$obj | Add-Member NoteProperty -Name "PublicFolderDistributionEnabled" -Value $Item.PublicFolderDistributionEnabled -Force
	$obj | Add-Member NoteProperty -Name "WebDistributionEnabled" -Value $Item.WebDistributionEnabled -Force
	$obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name -Force

	$Report += $OBJ
	}
	Return $Report
}
#endregion  OfflineAddressbook

#region Retention Policy
Function get-RetPolicy
{
#Add-text "Retention policy"
$ErrorActionPreference = "Stop"
$Report = @()

	If ($ShellVersion -eq "8")
	{
	 
	 $obj = New-Object PSObject
	 $obj | Add-Member NoteProperty -Name "RetentionPolicyTagLinks" -Value "N/A"
     $obj | Add-Member NoteProperty -Name "Name" -Value "N/A"
	 $Report += $obj
	}
	Else
	{
	Try{
    $Items = @(Get-RetentionPolicy)
   }
   Catch{
   #Add-text $_.message
   }
   Foreach ($Item in $Items)
   {
		   Try {
		     $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "RetentionPolicyTagLinks" -Value ($Item.RetentionPolicyTagLinks -join ", ")
		     $obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name
		     $Report += $obj
		      }
   Catch{ 
   #Add-text $_.Message
   }
   }
   }
	Return $Report
}
#endregion Retention Policy

#region RetentionTags
Function get-RetPolicyTag
{
#Add-text "Retention Tags"
$ErrorActionPreference = "Stop"
$Report = @()

	If ($ShellVersion -eq "8")
	{
	$obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "MessageClassDisplayName" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "MessageClass" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "RetentionEnabled" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "RetentionAction" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "AgeLimitForRetention" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "MoveToDestinationFolder" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "TriggerForRetention" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "Type" -Value "N/A"
			 $obj | Add-Member NoteProperty -Name "Identity" -Value "N/A"
	 $Report += $obj
	}
	Else
	{
	Try{
    $Items = @(Get-RetentionPolicyTag)
   }
   Catch{
   #Add-text $_.message
   }
   Foreach ($Item in $Items)
   {
		   Try {
		     $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "MessageClassDisplayName" -Value $Item.MessageClassDisplayName
		     $obj | Add-Member NoteProperty -Name "MessageClass" -Value $Item.MessageClass
			 $obj | Add-Member NoteProperty -Name "RetentionEnabled" -Value $Item.RetentionEnabled
			 $obj | Add-Member NoteProperty -Name "RetentionAction" -Value $Item.RetentionAction
			 $obj | Add-Member NoteProperty -Name "AgeLimitForRetention" -Value $Item.AgeLimitForRetention
			 $obj | Add-Member NoteProperty -Name "MoveToDestinationFolder" -Value $Item.MoveToDestinationFolder
			 $obj | Add-Member NoteProperty -Name "TriggerForRetention" -Value $Item.TriggerForRetention
		     $obj | Add-Member NoteProperty -Name "Type" -Value $Item.Type
			 $obj | Add-Member NoteProperty -Name "Identity" -Value $Item.Identity
			 $Report += $obj
		      }
   Catch{ 
   #Add-text $_.Message
   }
   }
   }
	Return $Report
}

#endregion RetentionTags

#region ABPolicies
Function get-ABPolicy
{
#Add-text "Addressbook Policy"
$ErrorActionPreference = "Stop"
$Report = @()
    
	If ($ShellVersion -eq "8")
	{
	 $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "AddressLists" -Value "N/A"
		     $obj | Add-Member NoteProperty -Name "GlobalAddressList" -Value  "N/A"
			 $obj | Add-Member NoteProperty -Name "RoomList" -Value  "N/A"
			 $obj | Add-Member NoteProperty -Name "OfflineAddressBook" -Value  "N/A"
			 $obj | Add-Member NoteProperty -Name "Name" -Value  "N/A"
	 $Report += $obj
	}
	Else
	{
	Try{
    $Items = @(Get-AddressBookPolicy)
   }
   Catch{#Add-text $_.message 
   }
   Foreach ($Item in $Items)
   {
		   Try {
		     $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "AddressLists" -Value ($Item.AddressLists -join ", ")
		     $obj | Add-Member NoteProperty -Name "GlobalAddressList" -Value $Item.GlobalAddressList
			 $obj | Add-Member NoteProperty -Name "RoomList" -Value $Item.RoomList
			 $obj | Add-Member NoteProperty -Name "OfflineAddressBook" -Value $Item.OfflineAddressBook
			 $obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name
		     $Report += $obj
		      }
   Catch{ 
   #Add-text $_.Message
   }
   }
   }
	Return $Report
}
#endregion ABPolicies

#region OWApolicy
Function get-OWAPolicy
{
#Add-text "OWA Policy"
$ErrorActionPreference = "Stop"
$Report = @()
    
	If ($ShellVersion -eq "8")
	{
	 $obj = New-Object PSObject
		     $obj | Add-Member NoteProperty -Name "Exchange 2007" -Value "OWA Policies N/A"
	 $Report += $obj
	 Return $Report
	}
	Else
	{
	Try{
    $Items = @(
	Get-OwaMailboxPolicy | Select DirectFileAccessOnPublicComputersEnabled,DirectFileAccessOnPrivateComputersEnabled,WebReadyDocumentViewingOnPublicComputersEnabled,WebReadyDocumentViewingOnPrivateComputersEnabled,`
	ForceWebReadyDocumentViewingFirstOnPublicComputers,ForceWebReadyDocumentViewingFirstOnPrivateComputers,ActionForUnknownFileAndMIMETypes:ForceSave,WebReadyDocumentViewingForAllSupportedTypes,`
	PhoneticSupportEnabled,DefaultTheme,DefaultClientLanguage,LogonAndErrorLanguage,UseGB18030,UseISO885915,OutboundCharset,GlobalAddressListEnabled,OrganizationEnabled,ExplicitLogonEnabled,OWALightEnabled,`
	OWAMiniEnabled,DelegateAccessEnabled,IRMEnabled,CalendarEnabled,ContactsEnabled,TasksEnabled,JournalEnabled,NotesEnabled,RemindersAndNotificationsEnabled,PremiumClientEnabled,SpellCheckerEnabled,`
	SearchFoldersEnabled,SignaturesEnabled,ThemeSelectionEnabled,JunkEmailEnabled,UMIntegrationEnabled,WSSAccessOnPublicComputersEnabled,WSSAccessOnPrivateComputersEnabled,ChangePasswordEnabled,`
	UNCAccessOnPublicComputersEnabled,UNCAccessOnPrivateComputersEnabled,ActiveSyncIntegrationEnabled,AllAddressListsEnabled,RulesEnabled,PublicFoldersEnabled,SMimeEnabled,RecoverDeletedItemsEnabled,`
	InstantMessagingEnabled,TextMessagingEnabled,ForceSaveAttachmentFilteringEnabled,SilverlightEnabled,InstantMessagingType,Name
	)
   }
   Catch{#Add-text $_.message 
   }
   Return Format-Array $Items
   
   }
 }
 <#
DirectFileAccessOnPublicComputersEnabled,DirectFileAccessOnPrivateComputersEnabled,WebReadyDocumentViewingOnPublicComputersEnabled,WebReadyDocumentViewingOnPrivateComputersEnabled,
ForceWebReadyDocumentViewingFirstOnPublicComputers,ForceWebReadyDocumentViewingFirstOnPrivateComputers,ActionForUnknownFileAndMIMETypes:ForceSave,WebReadyDocumentViewingForAllSupportedTypes,
PhoneticSupportEnabled,DefaultTheme,DefaultClientLanguage,LogonAndErrorLanguage,UseGB18030,UseISO885915,OutboundCharset,GlobalAddressListEnabled,OrganizationEnabled,ExplicitLogonEnabled,OWALightEnabled,
OWAMiniEnabled,DelegateAccessEnabled,IRMEnabled,CalendarEnabled,ContactsEnabled,TasksEnabled,JournalEnabled,NotesEnabled,RemindersAndNotificationsEnabled,PremiumClientEnabled,SpellCheckerEnabled,
SearchFoldersEnabled,SignaturesEnabled,ThemeSelectionEnabled,JunkEmailEnabled,UMIntegrationEnabled,WSSAccessOnPublicComputersEnabled,WSSAccessOnPrivateComputersEnabled,ChangePasswordEnabled,
UNCAccessOnPublicComputersEnabled,UNCAccessOnPrivateComputersEnabled,ActiveSyncIntegrationEnabled,AllAddressListsEnabled,RulesEnabled,PublicFoldersEnabled,SMimeEnabled,RecoverDeletedItemsEnabled,
InstantMessagingEnabled,TextMessagingEnabled,ForceSaveAttachmentFilteringEnabled,SilverlightEnabled,InstantMessagingType,Name
#>
#endregion OWApolicy

#region Remote Domain
Function GetremoteDomain
{
#Add-text "Remote Domains"
$Report = @()
$Remote = @(Get-RemoteDomain)
	Foreach ($item in $Remote)
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "DomainName" -Value $Item.DomainName -Force
	$obj | Add-Member NoteProperty -Name "CharacterSet" -Value $Item.CharacterSet -Force
	$obj | Add-Member NoteProperty -Name "NonMimeCharacterSet" -Value $Item.NonMimeCharacterSet -Force
	$obj | Add-Member NoteProperty -Name "AllowedOOFType" -Value $Item.AllowedOOFType -Force
	$obj | Add-Member NoteProperty -Name "AutoReplyEnabled" -Value $Item.AutoReplyEnabled -Force
	$obj | Add-Member NoteProperty -Name "AutoForwardEnabled" -Value $Item.AutoForwardEnabled -Force
	$obj | Add-Member NoteProperty -Name "DeliveryReportEnabled" -Value $Item.DeliveryReportEnabled -Force
	$obj | Add-Member NoteProperty -Name "NDREnabled" -Value $Item.NDREnabled -Force
	$obj | Add-Member NoteProperty -Name "MeetingForwardNotificationEnabled" -Value $Item.MeetingForwardNotificationEnabled -Force
	$obj | Add-Member NoteProperty -Name "ContentType" -Value $Item.ContentType -Force
	$obj | Add-Member NoteProperty -Name "DisplaySenderName" -Value $Item.DisplaySenderName -Force
	$obj | Add-Member NoteProperty -Name "TNEFEnabled" -Value $Item.TNEFEnabled -Force
	$obj | Add-Member NoteProperty -Name "LineWrapSize" -Value $Item.LineWrapSize -Force
	$obj | Add-Member NoteProperty -Name "UseSimpleDisplayName" -Value $Item.UseSimpleDisplayName -Force
	$obj | Add-Member NoteProperty -Name "Identity" -Value $Item.Identity -Force

	$Report += $OBJ
	}
	Return $Report
}

#endregion Remote Domain

#region ActiveSyncPolicy
Function get-ActiveSyncPolicy
{
#Add-text "ActiveSync"
$ErrorActionPreference = "SilentlyContinue"
$Report = @()
    
	If ((@(Get-Command Get-MobileDeviceMailboxPolicy).count) -gt 0 )
	{
	    $Items = @(
		Get-MobileDeviceMailboxPolicy | select AllowNonProvisionableDevices,AlphanumericPasswordRequired,AttachmentsEnabled,DeviceEncryptionEnabled,RequireStorageCardEncryption,PasswordEnabled,PasswordRecoveryEnabled,DevicePolicyRefreshInterval,AllowSimplePassword,MaxAttachmentSize,`
WSSAccessEnabled,UNCAccessEnabled,MinPasswordLength,MaxInactivityTimeLock,MaxPasswordFailedAttempts,PasswordExpiration,PasswordHistory,IsDefault,AllowApplePushNotifications,AllowMicrosoftPushNotifications,AllowStorageCard,AllowCamera,`
RequireDeviceEncryption,AllowUnsignedApplications,AllowUnsignedInstallationPackages,AllowWiFi,AllowTextMessaging,AllowPOPIMAPEmail,AllowIrDA,RequireManualSyncWhenRoaming,AllowDesktopSync,AllowHTMLEmail,RequireSignedSMIMEMessages,RequireEncryptedSMIMEMessages,`
AllowSMIMESoftCerts,AllowBrowser,AllowConsumerEmail,AllowRemoteDesktop,AllowInternetSharing,AllowBluetooth,MaxCalendarAgeFilter,MaxEmailAgeFilter,RequireSignedSMIMEAlgorithm,RequireEncryptionSMIMEAlgorithm,AllowSMIMEEncryptionAlgorithmNegotiation,`
MinPasswordComplexCharacters,MaxEmailBodyTruncationSize,MaxEmailHTMLBodyTruncationSize,UnapprovedInROMApplicationList,ApprovedApplicationList,AllowExternalDeviceManagement,MobileOTAUpdateMode,AllowMobileOTAUpdate,IrmEnabled,Name
		)
	 Return Format-Array $Items
	}
	Else
	{

    $Items = @(
	Get-ActiveSyncMailboxPolicy | select AllowNonProvisionableDevices,AlphanumericPasswordRequired,AttachmentsEnabled,DeviceEncryptionEnabled,RequireStorageCardEncryption,PasswordEnabled,PasswordRecoveryEnabled,DevicePolicyRefreshInterval,AllowSimplePassword,MaxAttachmentSize,`
WSSAccessEnabled,UNCAccessEnabled,MinPasswordLength,MaxInactivityTimeLock,MaxPasswordFailedAttempts,PasswordExpiration,PasswordHistory,IsDefault,AllowApplePushNotifications,AllowMicrosoftPushNotifications,AllowStorageCard,AllowCamera,`
RequireDeviceEncryption,AllowUnsignedApplications,AllowUnsignedInstallationPackages,AllowWiFi,AllowTextMessaging,AllowPOPIMAPEmail,AllowIrDA,RequireManualSyncWhenRoaming,AllowDesktopSync,AllowHTMLEmail,RequireSignedSMIMEMessages,RequireEncryptedSMIMEMessages,`
AllowSMIMESoftCerts,AllowBrowser,AllowConsumerEmail,AllowRemoteDesktop,AllowInternetSharing,AllowBluetooth,MaxCalendarAgeFilter,MaxEmailAgeFilter,RequireSignedSMIMEAlgorithm,RequireEncryptionSMIMEAlgorithm,AllowSMIMEEncryptionAlgorithmNegotiation,`
MinPasswordComplexCharacters,MaxEmailBodyTruncationSize,MaxEmailHTMLBodyTruncationSize,UnapprovedInROMApplicationList,ApprovedApplicationList,AllowExternalDeviceManagement,MobileOTAUpdateMode,AllowMobileOTAUpdate,IrmEnabled,Name
	)

   Return Format-Array $Items
   
   }
 }
 #endregion ActiveSyncPolicy
 
#region TransportConfig
Function  GetTransportConfig
{
#Add-text "Transport Config"
$Report = @()
$Obj = Get-TransportConfig | select *
Add-Member –InputObject $Obj -MemberType NoteProperty -Name Name –Value "Transport Config" -Force
$Report += $OBJ

	Return Format-Array $Report
}
#endregion TransportConfig
 
#region DialPlan

Function GetUMDialplan 
{
#Add-text "UM Dial Plan"
$Report = @(
Get-UMDialPlan | select NumberOfDigitsInExtension,LogonFailuresBeforeDisconnect,AccessTelephoneNumbers,FaxEnabled,InputFailuresBeforeDisconnect,OutsideLineAccessCode,DialByNamePrimary,DialByNameSecondary,AudioCodec,DefaultLanguage,VoIPSecurity,`
			MaxCallDuration,MaxRecordingDuration,RecordingIdleTimeout,PilotIdentifierList,UMServers,UMMailboxPolicies,UMAutoAttendants,WelcomeGreetingEnabled,AutomaticSpeechRecognitionEnabled,PhoneContext,WelcomeGreetingFilename,`
			InfoAnnouncementFilename,OperatorExtension,DefaultOutboundCallingLineId,Extension,MatchedNameSelectionMethod,InfoAnnouncementEnabled,InternationalAccessCode,NationalNumberPrefix,InCountryOrRegionNumberFormat,InternationalNumberFormat,`
			CallSomeoneEnabled,ContactScope,ContactAddressList,SendVoiceMsgEnabled,UMAutoAttendant,AllowDialPlanSubscribers,AllowExtensions,AllowedInCountryOrRegionGroups,AllowedInternationalGroups,ConfiguredInCountryOrRegionGroups,LegacyPromptPublishingPoint,`
			ConfiguredInternationalGroups,UMIPGateway,URIType,SubscriberType,GlobalCallRoutingScheme,TUIPromptEditingEnabled,CallAnsweringRulesEnabled,SipResourceIdentifierRequired,FDSPollingInterval,EquivalentDialPlanPhoneContexts,NumberingPlanFormats,`
			AllowHeuristicADCallingLineIdResolution,CountryOrRegionCode,AdminDisplayName,ExchangeVersion,Name
)
Return Format-Array $Report
}

#endregion DialPlan

#region UMgateway

Function getumIPGateway 
{
#Add-text "UM gateway"
$Report = @()
$Gateway = @(Get-UMIPGateway)
	Foreach ($item in $Gateway)
	{
	$OBJ = New-Object PSObject
	$obj | Add-Member NoteProperty -Name "Address" -Value $Item.Address -Force
	$obj | Add-Member NoteProperty -Name "OutcallsAllowed" -Value $Item.OutcallsAllowed -Force
	$obj | Add-Member NoteProperty -Name "Status" -Value $Item.Status -Force
	$obj | Add-Member NoteProperty -Name "Port" -Value $Item.Port -Force
	$obj | Add-Member NoteProperty -Name "Simulator" -Value $Item.Simulator -Force
	$obj | Add-Member NoteProperty -Name "IPAddressFamily" -Value $Item.IPAddressFamily -Force
	$obj | Add-Member NoteProperty -Name "DelayedSourcePartyInfoEnabled" -Value $Item.DelayedSourcePartyInfoEnabled -Force
	$obj | Add-Member NoteProperty -Name "MessageWaitingIndicatorAllowed" -Value $Item.MessageWaitingIndicatorAllowed -Force
	$obj | Add-Member NoteProperty -Name "HuntGroups" -Value ($Item.HuntGroups -join ", ") -Force
	$obj | Add-Member NoteProperty -Name "GlobalCallRoutingScheme" -Value $Item.GlobalCallRoutingScheme -Force
	$obj | Add-Member NoteProperty -Name "ForwardingAddress" -Value $Item.ForwardingAddress -Force
	$obj | Add-Member NoteProperty -Name "Name" -Value $Item.Name -Force
	$Report += $OBJ
	}
	Return $Report
}

#endregion UMgateway

#region UMPolicy

Function GetUMpolicy 
{
$Report = @(
#Add-text "UM Policy"
Get-UMMailboxPolicy | select MaxGreetingDuration,MaxLogonAttempts,AllowCommonPatterns,PINLifetime,PINHistoryCount,AllowSMSNotification,ProtectUnauthenticatedVoiceMail,ProtectAuthenticatedVoiceMail,ProtectedVoiceMailText,RequireProtectedPlayOnPhone,`
MinPINLength,FaxMessageText,UMEnabledText,ResetPINText,SourceForestPolicyNames,VoiceMailText,UMDialPlan,FaxServerURI,AllowedInCountryOrRegionGroups,AllowedInternationalGroups,AllowDialPlanSubscribers,AllowExtensions,`
LogonFailuresBeforePINReset,AllowMissedCallNotifications,AllowFax,AllowTUIAccessToCalendar,AllowTUIAccessToEmail,AllowSubscriberAccess,AllowTUIAccessToDirectory,AllowTUIAccessToPersonalContacts,AllowAutomaticSpeechRecognition,`
AllowPlayOnPhone,AllowVoiceMailPreview,AllowCallAnsweringRules,AllowMessageWaitingIndicator,AllowPinlessVoiceMailAccess,AllowVoiceResponseToOtherMessageTypes,AllowVoiceMailAnalysis,AllowVoiceNotification,`
InformCallerOfVoiceMailAnalysis,VoiceMailPreviewPartnerAddress,VoiceMailPreviewPartnerAssignedID,VoiceMailPreviewPartnerMaxMessageDuration,VoiceMailPreviewPartnerMaxDeliveryDelay,IsDefault,Name
)
Return Format-Array $Report
}

#endregion UMPolicy
 
#region UMAutoAttendant

Function GetUMAutoAttendant
{
#Add-text "UM Attendant"
$Report = @(
Get-UMAutoAttendant | select SpeechEnabled,AllowDialPlanSubscribers,AllowExtensions,AllowedInCountryOrRegionGroups,AllowedInternationalGroups,CallSomeoneEnabled,ContactScope,ContactAddressList,SendVoiceMsgEnabled,PilotIdentifierList,`
UMDialPlan,DTMFFallbackAutoAttendant,HolidaySchedule,TimeZone,TimeZoneName,MatchedNameSelectionMethod,BusinessLocation,WeekStartDay,Status,Language,OperatorExtension,InfoAnnouncementFilename,InfoAnnouncementEnabled,`
NameLookupEnabled,StarOutToDialPlanEnabled,ForwardCallsToDefaultMailbox,DefaultMailbox,BusinessName,BusinessHoursWelcomeGreetingFilename,BusinessHoursWelcomeGreetingEnabled,BusinessHoursMainMenuCustomPromptFilename,`
BusinessHoursMainMenuCustomPromptEnabled,BusinessHoursTransferToOperatorEnabled,BusinessHoursKeyMapping,BusinessHoursKeyMappingEnabled,AfterHoursWelcomeGreetingFilename,AfterHoursWelcomeGreetingEnabled,`
AfterHoursMainMenuCustomPromptFilename,AfterHoursMainMenuCustomPromptEnabled,AfterHoursTransferToOperatorEnabled,AfterHoursKeyMapping,AfterHoursKeyMappingEnabled,Name
)
Return Format-Array $Report
}

#endregion UMAutoAttendant

#region EmailaddressPolicy

Function GetEmailaddresspolicy
{
#Add-text "Email Address Policy"
$Report = @(
Get-EmailAddressPolicy | select RecipientFilter,LdapRecipientFilter,LastUpdatedRecipientFilter,RecipientFilterApplied,IncludedRecipients,ConditionalDepartment,ConditionalCompany,ConditionalStateOrProvince,`
ConditionalCustomAttribute1,ConditionalCustomAttribute2,ConditionalCustomAttribute3,ConditionalCustomAttribute4,ConditionalCustomAttribute5,ConditionalCustomAttribute6,`
ConditionalCustomAttribute7,ConditionalCustomAttribute8,ConditionalCustomAttribute9,ConditionalCustomAttribute10,ConditionalCustomAttribute11,ConditionalCustomAttribute12,`
ConditionalCustomAttribute13,ConditionalCustomAttribute14,ConditionalCustomAttribute15,RecipientContainer,RecipientFilterType,Priority,EnabledPrimarySMTPAddressTemplate,`
EnabledEmailAddressTemplates,DisabledEmailAddressTemplates,HasEmailAddressSetting,HasMailboxManagerSetting,NonAuthoritativeDomains,ExchangeVersion,Name
)
Return Format-Array $Report
}

Function Get-DagNetwork
{
#Add-text "DAG Networks"
try{
$report = @(
Get-DatabaseAvailabilityGroupNetwork | select Identity,@{N='Subnets';E={$_.Subnets -join [char]10}} ,@{N='Interfaces';E={$_.Interfaces -join [Char]10}},MapiAccessEnabled,ReplicationEnabled,IgnoreNetwork -ErrorAction SilentlyContinue
)
}
Catch {}
Return $report
}
#endregion EmailaddressPolicy
 
#region RBAC
Function Get-RBAC
{
#Add-text "RBAC"
$List = Get-ManagementRoleAssignment -GetEffectiveUsers
$AdminRoles = $List | Select RoleAssigneeName -ExpandProperty RoleAssigneeName | Select -Unique
$AdminRight = @{}
#$Members = @{}
$Accounts = @{}
ForEach ($Gr in $AdminRoles)
{
$AdminRight.$gr = ($List | ?{$_.RoleAssigneeName -eq $Gr} | Select Role -Unique)

}

$UsrPath = $user.User -split '/'

Function Build-DN
{
    Param([int]$str,[Array]$UsrPath)
    IF ($str -eq -1)
    {
    Return [string]"CN=$($UsrPath[$str]),"
    }
    ELSEIF ($str -eq [int]"-$($UsrPath.count)")
    {      
                IF(((($UsrPath[$str]) -split "\.").count) -eq 4) 
                {
                Return [string]"DC=$(((($UsrPath[$str]) -split "\.")[0])),DC=$((($UsrPath[$str]) -split "\.")[1]),DC=$((($UsrPath[$str]) -split "\.")[2]),DC=$((($UsrPath[$str]) -split "\.")[3])"
                }
                ELSEIF(((($UsrPath[$str]) -split "\.").count) -eq 3) 
                {
                Return [string]"DC=$(((($UsrPath[$str]) -split "\.")[0])),DC=$((($UsrPath[$str]) -split "\.")[1]),DC=$((($UsrPath[$str]) -split "\.")[2])"
                }
                ELSEIF (((($UsrPath[$str]) -split "\.").count) -eq 2) 
                {
                Return [string]"DC=$(((($UsrPath[$str]) -split "\.")[0])),DC=$((($UsrPath[$str]) -split "\.")[1])"
                }

    }
    ELSEIF ($str -eq ([int]"-$($UsrPath.count)" + 1) -and $UsrPath[$([int]"-$($UsrPath.count)" + 1)] -eq 'Users')
    {
    Return [String]"CN=$($UsrPath[$str]),"
    }
    ELSE
    {
    Return [String]"OU=$($UsrPath[$str]),"
    }
}


ForEach ($AC in @($List | Select User -Unique))
{
$Usr = $AC.User -split "/"

Write-Verbose "Finding Object: $($AC.User) = $($Usr.Count)"
$DistinguishedName = ""
-1..-[int]"$($Usr.Count)" | %{$DistinguishedName = $DistinguishedName + "$(Build-DN -str $_ -UsrPath $Usr)" }
Write-Verbose $DistinguishedName 
try{
$Accounts."$($AC.user)" = Get-ADObject -Identity $DistinguishedName -Properties DisplayName -Server "$(($Usr)[0])"  -ErrorAction SilentlyContinue
}
Catch{}
}
$Report = @()
ForEach ($row in $AdminRoles)
{

$members = @()
$list | ?{$_.RoleAssigneeName -eq $row} | Select User -Unique | %{$members += ($Accounts["$($_.user)"].DisplayName) }
$Values = @{
AdminRole = $Row
Role      = @($AdminRight.$row.Role) -join "`n"
Members   = @($members | ?{$_ -ne $null}) -join "`n"

          }
$obj = New-Object -TypeName PSCustomObject -Property $Values
$Report += $obj 

}
Return ($Report | Select AdminRole,Role,Members -Unique)
}

#endregion

#region DagDistribution
Function Get-DAGDistribution
{
#Add-text "DAG Distribution Preference"
    $info = @()
    $DagServers = @(Get-DatabaseAvailabilityGroup | select Name,Servers -Unique).Servers
    $DagServers | %{$info += Get-MailboxDatabaseCopyStatus -Server "$($_)" | Select DatabaseName,ActiveDatabaseCopy,ActivationPreference,MailboxServer} 
    $DBNames = @($info | Select DatabaseName -Unique) 

    $ObjectCollection = @()
    #}
            Foreach ($DB in $DBNames)
            {
                    $obj = New-Object -TypeName PSCustomObject
                   # $obj | Add-Member -MemberType NoteProperty -Name 'DAG' -Value "$($DAG.Name)"
                    $obj | Add-Member -MemberType NoteProperty -Name 'Database' -Value "$($DB.DatabaseName)"
                
                    Foreach ($svr in @($DagServers ))
                    {
                    $AP = ($info | ?{$_.MailboxServer -eq "$($svr)" -and $_.DatabaseName -eq "$($DB.DatabaseName)"}).ActivationPreference
                    $obj | Add-Member -MemberType NoteProperty -Name "$($svr)" -Value $AP 
 
                    }
                    $ObjectCollection += $obj
            }

     
   # }
    $ObjectCollection | Export-Csv -Path ".\data\DAGDistribution_DAG.CSV"
}


#endregion


#region hybrid
Function get-Hybrid
{
try{
#Add-text "O365 Hybrid Config"
Get-HybridConfiguration | select @{N='ReceivingTransportServers';E={$_.ReceivingTransportServers -join "`n"}},OnPremisesSmartHost,`
@{N='SendingTransportServers';E={$_.SendingTransportServers -join "`n"}},`
@{N='Domains';E={$_.Domains -join "`n"}},`
@{N='Features';E={$_.Features -join "`n"}}
}
catch{}
}

#endregion

#region MSOL
<#
Function Get-Sku

{
try{
Get-MsolAccountSku | select SkuPartNumber,ActiveUnits,ConsumedUnits,SuspendedUnits,WarningUnits
}
catch{}
}


Function Get-MsolCompInfo
{
#Add-text "O365 CompanyInfo"
Get-MsolCompanyInformation | select DisplayName,PreferredLanguage,TelephoneNumber,`
@{n='TechnicalNotificationEmails';e={$_.TechnicalNotificationEmails -join "`n"}},`
SelfServePasswordResetEnabled,UsersPermissionToCreateGroupsEnabled,UsersPermissionToCreateLOBAppsEnabled ,`
UsersPermissionToReadOtherUsersEnabled,UsersPermissionToUserConsentToAppEnabled,DirectorySynchronizationEnabled,PasswordSynchronizationEnabled

}

Function get-MsolD
{
try{
#Add-text "O365 Domains"
Get-MsolDomain | Select Authentication,Capabilities,IsDefault,IsInitial,Name,`
@{N='Status';E={"$($_.Status) Method:$($_.VerificationMethod)" }}
}
Catch{}
}
#>

Function Get-O365Orgrelationship
{
try{
Get-O365OrganizationRelationship | select @{Name='DomainNames';E={$_.DomainNames -join ","}},Free*,Mailtips*,Target*,Name
}
Catch
{}
}


Function Get-O365AcceptedDom
{
try{
Get-O365AcceptedDomain | select domainname, domaintype, default
}
Catch
{}
}


Function Get-O365RemotedDom
{
try{
Get-O365RemoteDomain  | select DomainName, TargetDeliveryDomain 
}
Catch
{}
}


Function get-IntraOrgCon
{
try{
Get-O365IntraOrganizationConnector | select @{n='TargetAddressDomains';e={$_.TargetAddressDomains -join "`n"}},`
DiscoveryEndpoint,Enabled
}
Catch{}
}


Function Get-O365IConnector
{
try{
#Add-text "O365 Connectors 1"
Get-O365InboundConnector | Select Enabled,@{n='SenderIPAddresses';e={$_.SenderIPAddresses -join "`n"}},`
@{n='SenderDomains';e={$_.SenderDomains -join "`n"}},`
@{n='AssociatedAcceptedDomains';e={$_.AssociatedAcceptedDomains -join "`n"}},`
RequireTls,Name,ConnectorType
}
Catch{}
}

Function Get-O365OConnector
{
try{
#Add-text "O365 Connectors 2"
Get-O365OutboundConnector | Select Enabled,@{n='RecipientDomains';e={$_.RecipientDomains -join "`n"}},`
@{n='SmartHosts';e={$_.SmartHosts -join "`n"}},TlsDomain,Name,ConnectorType
}
Catch{}
}


Function Get-O365Filter
{
try{
#Add-text "O365 Filters"
Get-O365HostedConnectionFilterPolicy | select IsDefault,
@{N='IPAllowList';E={$_.IPAllowList -join "`n"}},`
@{N='IPBlockList';E={$_.IPBlockList -join "`n"}},`
@{N='Name';E={$_.Identity }}
}
Catch{}
}


Function Get-O365TrRule
{
try{
#Add-text "O365 Rules"
Get-O365TransportRule | select "Identity","State","Comments","Description"
}
Catch{}
}

Function get-O365fed
{
try{
#Add-text "O365 Federation Information"
Get-O365FederationTrust | select ApplicationUri,Identity
}
Catch{}
}

Function get-O365Sharing
{
try{
#Add-text "O365 Sharing Policy"
Get-O365SharingPolicy | select Enabled,Default,ID,@{N='Shared Domains';E={$_.Domains -join "`n"}}
}
Catch{}
}


#endregion

function get-availibiltyAdd
{
try{
Get-AvailabilityAddressSpace | select ForestName,UserName,UseServiceAccount,AccessMethod,ProxyUrl,TargetAutodiscoverEpr,Name
}
catch
{}

}