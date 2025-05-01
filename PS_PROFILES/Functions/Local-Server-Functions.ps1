
## Begin Get-LocalAdmin
Function Get-LocalAdmin {
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    # Get local administrators
    $admins = Get-WmiObject -Class Win32_GroupUser -ComputerName $ComputerName |
              Where-Object { $_.GroupComponent -like '*"Administrators"' }

    # Extract and format user names
    $admins | ForEach-Object {
        if ($_.PartComponent -match 'Domain\=(.+?), Name\=(.+)$') {
            "$($matches[1].Trim('"'))\$($matches[2].Trim('"'))"
        }
    }
}
## End Get-LocalAdmin
## Begin Get-LocalAdministratorBuiltin
Function Get-LocalAdministratorBuiltin{
<#
	.SYNOPSIS
		Function to retrieve the local Administrator account
	
	.DESCRIPTION
		Function to retrieve the local Administrator account
	
	.PARAMETER ComputerName
		Specifies the computername
	
	.EXAMPLE
		PS C:\> Get-LocalAdministratorBuiltin
	
	.EXAMPLE
		PS C:\> Get-LocalAdministratorBuiltin -ComputerName SERVER01
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
	#Function to get the BUILTIN LocalAdministrator
	#http://blog.simonw.se/powershell-find-builtin-local-administrator-account/
#>
	
	[CmdletBinding()]
	param (
		[Parameter()]
		$ComputerName = $env:computername
	)
	Process
	{
		Foreach ($Computer in $ComputerName)
		{
			Try
			{
				Add-Type -AssemblyName System.DirectoryServices.AccountManagement
				$PrincipalContext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine, $Computer)
				$UserPrincipal = New-Object -TypeName System.DirectoryServices.AccountManagement.UserPrincipal($PrincipalContext)
				$Searcher = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalSearcher
				$Searcher.QueryFilter = $UserPrincipal
				$Searcher.FindAll() | Where-Object { $_.Sid -Like "*-500" }
			}
			Catch
			{
				Write-Warning -Message "$($_.Exception.Message)"
			}
		}
	}
}
## End Get-LocalAdministratorBuiltin
## Begin Get-LocalGroup
Function Get-LocalGroup{
	
<#
	.SYNOPSIS
		This script can be list all of local group account.
	
	.DESCRIPTION
		This script can be list all of local group account.
		The Function is using WMI to connect to the remote machine
	
	.PARAMETER ComputerName
		Specifies the computers on which the command . The default is the local computer.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	
	.EXAMPLE
		Get-LocalGroup
		
		This example shows how to list all the local groups on local computer.
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	PARAM
	(
		[Alias('cn')]
		[String[]]$ComputerName = $Env:COMPUTERNAME,
		
		[String]$AccountName,
		
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$Splatting = @{
		Class = "Win32_Group"
		Namespace = "root\cimv2"
		Filter = "LocalAccount='$True'"
	}
	
	#Credentials
	If ($PSBoundParameters['Credential']) { $Splatting.Credential = $Credential }
	
	Foreach ($Computer in $ComputerName)
	{
		TRY
		{
			Write-Verbose -Message "[PROCESS] ComputerName: $Computer"
			Get-WmiObject @Splatting -ComputerName $Computer | Select-Object -Property Name, Caption, Status, SID, SIDType, Domain, Description
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Issue connecting to $Computer"
		}
	}
}
## End Get-LocalGroup
## Begin Get-LocalGroupMember
Function Get-LocalGroupMember {

<#
.SYNOPSIS
Get local group membership using ADSI.

.DESCRIPTION
This command uses ADSI to connect to a server and enumerate the members of a local group. By default it will retrieve members of the local Administrators group.

The command uses legacy protocols to connect and enumerate group memberships. You may find it more efficient to wrap this Function in an Invoke-Command expression. See examples.

.PARAMETER Computername
The name of a computer to query. The parameter has aliases of 'CN' and 'Host'.

.PARAMETER Name
The name of a local group. 

.EXAMPLE
PS C:\> Get-LocalGroupMember -computer chi-core01

Computername : CHI-CORE01
Name         : Administrator
ADSPath      : WinNT://GLOBOMANTICS/chi-core01/Administrator
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : True

Computername : CHI-CORE01
Name         : Domain Admins
ADSPath      : WinNT://GLOBOMANTICS/Domain Admins
Class        : Group
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : Chicago IT
ADSPath      : WinNT://GLOBOMANTICS/Chicago IT
Class        : Group
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : OMAA
ADSPath      : WinNT://GLOBOMANTICS/OMAA
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-CORE01
Name         : LocalAdmin
ADSPath      : WinNT://GLOBOMANTICS/chi-core01/LocalAdmin
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : True

.EXAMPLE
PS C:\> "chi-hvr1","chi-hvr2","chi-core01","chi-fp02" | get-localgroupmember  | where {$_.IsLocal} | Select Computername,Name,ADSPath

Computername Name          ADSPath                                      
------------ ----          -------                                      
CHI-HVR1     Administrator WinNT://GLOBOMANTICS/chi-hvr1/Administrator  
CHI-HVR2     Administrator WinNT://GLOBOMANTICS/chi-hvr2/Administrator  
CHI-HVR2     Jeff          WinNT://GLOBOMANTICS/chi-hvr2/Jeff           
CHI-CORE01   Administrator WinNT://GLOBOMANTICS/chi-core01/Administrator
CHI-CORE01   LocalAdmin    WinNT://GLOBOMANTICS/chi-core01/LocalAdmin   
CHI-FP02     Administrator WinNT://GLOBOMANTICS/chi-fp02/Administrator

.EXAMPLE
PS C:\> $s = new-pssession chi-hvr1,chi-fp02,chi-hvr2,chi-core01
Create several PSSessions to remote computers.

PS C:\> $sb = ${Function:Get-localGroupMember} 

Get the Function's scriptblock

PS C:\> Invoke-Command -scriptblock { new-item -path Function:Get-LocalGroupMember -value $using:sb} -session $s 

Create a remote version of the Function.

PS C:\> Invoke-Command -scriptblock { get-localgroupmember | where {$_.IsLocal} } -session $s | Select Computername,Name,ADSPath

Repeat an example from above but this time execute it in a remote session.

.EXAMPLE
PS C:\> get-localgroupmember -Name "Hyper-V administrators" -Computername chi-hvr1,chi-hvr2


Computername : CHI-HVR1
Name         : jeff
ADSPath      : WinNT://GLOBOMANTICS/jeff
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Computername : CHI-HVR2
Name         : jeff
ADSPath      : WinNT://GLOBOMANTICS/jeff
Class        : User
Domain       : GLOBOMANTICS
IsLocal      : False

Check group membership for the Hyper-V Administrators group.

.EXAMPLE
PS C:\> get-localgroupmember -Computername chi-core01 | where class -eq 'group' | select Domain,Name

Domain       Name         
------       ----         
GLOBOMANTICS Domain Admins
GLOBOMANTICS Chicago IT   

Get members of the Administrators group on CHI-CORE01 that are groups and select a few properties.


.NOTES
NAME        :  Get-LocalGroupMember
VERSION     :  1.6   
LAST UPDATED:  2/18/2016
AUTHOR      :  Jeff Hicks (@JeffHicks)

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

.INPUTS
[string] for computer names

.OUTPUTS
[object]

#>


[cmdletbinding()]

Param(
[Parameter(Position = 0)]
[ValidateNotNullorEmpty()]
[string]$Name = "Administrators",

[Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
[ValidateNotNullorEmpty()]
[Alias("CN","host")]
[string[]]$Computername = $env:computername
)


Begin {
    Write-Verbose "[Starting] $($MyInvocation.Mycommand)"  
    Write-Verbose "[Begin]    Querying members of the $Name group"
} #begin
## End Get-LocalGroupMember

Process {
 
 foreach ($computer in $computername) {

    #define a flag to indicate if there was an error
    $script:NotFound = $False
    
    #define a trap to handle errors because we're not using cmdlets that
    #could support Try/Catch. Traps must be in same scope.
    Trap [System.Runtime.InteropServices.COMException] {
        $errMsg = "Failed to enumerate $name on $computer. $($_.exception.message)"
        Write-Warning $errMsg

        #set a flag
        $script:NotFound = $True
    
        Continue    
    }

    #define a Trap for all other errors
    Trap {
      Write-Warning "Oops. There was some other type of error: $($_.exception.message)"
      Continue
    }

    Write-Verbose "[Process]  Connecting to $computer"
    #the WinNT moniker is case-sensitive
    [ADSI]$group = "WinNT://$computer/$Name,group"
        
    Write-Verbose "[Process]  Getting group member details" 
    $members = $group.invoke("Members") 

    Write-Verbose "[Process]  Counting group members"
    
    if (-Not $script:NotFound) {
        $found = ($members | measure).count
        Write-Verbose "[Process]  Found $found members"

        if ($found -gt 0 ) {
        $members | foreach {
        
            #define an ordered hashtable which will hold properties
            #for a custom object
            $Hash = [ordered]@{Computername = $computer.toUpper()}

            #Get the name property
            $hash.Add("Name",$_[0].GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null))
        
            #get ADS Path of member
            $ADSPath = $_[0].GetType().InvokeMember("ADSPath", 'GetProperty', $null, $_, $null)
            $hash.Add("ADSPath",$ADSPath)
    
            #get the member class, ie user or group
            $hash.Add("Class",$_[0].GetType().InvokeMember("Class", 'GetProperty', $null, $_, $null))  
    
            <#
            Domain members will have an ADSPath like WinNT://MYDomain/Domain Users.  
            Local accounts will be like WinNT://MYDomain/Computername/Administrator
            #>

            $hash.Add("Domain",$ADSPath.Split("/")[2])

            #if computer name is found between two /, then assume
            #the ADSPath reflects a local object
            if ($ADSPath -match "/$computer/") {
                $local = $True
                }
            else {
                $local = $False
                }
            $hash.Add("IsLocal",$local)

            #turn the hashtable into an object
            New-Object -TypeName PSObject -Property $hash
         } #foreach member
        } 
        else {
            Write-Warning "No members found in $Name on $Computer."
        }
    } #if no errors
} #foreach computer
## End Get-LocalGroupMember

} #process
## End Get-LocalGroupMember

End {
    Write-Verbose "[Ending]  $($MyInvocation.Mycommand)"
} #end
## End Get-LocalGroupMember

} #end Function
## End Get-LocalGroupMember
## Begin Get-LocalGroupMembers
Function Get-LocalGroupMembers{
	<#
Copyright (c) 2016 JDH Information Technology Solutions, Inc.


Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:


The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.


THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
#>
<#
===================================================================================  
DESCRIPTION:    Function enumerates members of all local groups (or a given group). 
If -Server parameter is not specified, it will query localhost by default. 
If -Group parameter is not specified, all local groups will be queried. 
            
AUTHOR:    	Piotr Lewandowski 
VERSION:    1.0  
DATE:       29/04/2013  
SYNTAX:     Get-LocalGroupMembers [[-Server] <string[]>] [[-Group] <string[]>] 
             
EXAMPLES:   

Get-LocalGroupMembers -server "scsm-server" | ft -AutoSize

Server      Local Group          Name                 Type  Domain  SID
------      -----------          ----                 ----  ------  ---
scsm-server Administrators       Administrator        User          S-1-5-21-1473970658-40817565-21663372-500
scsm-server Administrators       Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
scsm-server Guests               Guest                User          S-1-5-21-1473970658-40817565-21663372-501
scsm-server Remote Desktop Users pladmin              User  contoso S-1-5-21-4081441239-4240563405-729182456-1272
scsm-server Users                INTERACTIVE          Group         S-1-5-4
scsm-server Users                Authenticated Users  Group         S-1-5-11



"scsm-dc01","scsm-server" | Get-LocalGroupMembers -group administrators | ft -autosize

Server      Local Group    Name                 Type  Domain  SID
------      -----------    ----                 ----  ------  ---
scsm-dc01   administrators Administrator        User  contoso S-1-5-21-4081441239-4240563405-729182456-500
scsm-dc01   administrators Enterprise Admins    Group contoso S-1-5-21-4081441239-4240563405-729182456-519
scsm-dc01   administrators Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
scsm-server administrators Administrator        User          S-1-5-21-1473970658-40817565-21663372-500
scsm-server administrators !svcServiceManager   User  contoso S-1-5-21-4081441239-4240563405-729182456-1274
scsm-server administrators !svcServiceManagerWF User  contoso S-1-5-21-4081441239-4240563405-729182456-1275
scsm-server administrators !svcscoservice       User  contoso S-1-5-21-4081441239-4240563405-729182456-1310
scsm-server administrators Domain Admins        Group contoso S-1-5-21-4081441239-4240563405-729182456-512
 
===================================================================================  

#>
param(
[Parameter(ValuefromPipeline=$true)][array]$server = $env:computername,
$GroupName = $null
)
PROCESS {
    $finalresult = @()
    $computer = [ADSI]"WinNT://$server"

    if (!($groupName))
    {
    $Groups = $computer.psbase.Children | Where {$_.psbase.schemaClassName -eq "group"} | select -expand name
    }
    else
    {
    $groups = $groupName
    }
    $CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().GetDirectoryEntry() | select name,objectsid
    $domain = $currentdomain.name
    $SID=$CurrentDomain.objectsid
    $DomainSID = (New-Object System.Security.Principal.SecurityIdentifier($sid[0], 0)).value


    foreach ($group in $groups)
    {
    $gmembers = $null
    $LocalGroup = [ADSI]("WinNT://$server/$group,group")


    $GMembers = $LocalGroup.psbase.invoke("Members")
    $GMemberProps = @{Server="$server";"Local Group"=$group;Name="";Type="";ADSPath="";Domain="";SID=""}
    $MemberResult = @()


        if ($gmembers)
        {
        foreach ($gmember in $gmembers)
            {
            $membertable = new-object psobject -Property $GMemberProps
            $name = $gmember.GetType().InvokeMember("Name",'GetProperty', $null, $gmember, $null)
            $sid = $gmember.GetType().InvokeMember("objectsid",'GetProperty', $null, $gmember, $null)
            $UserSid = New-Object System.Security.Principal.SecurityIdentifier($sid, 0)
            $class = $gmember.GetType().InvokeMember("Class",'GetProperty', $null, $gmember, $null)
            $ads = $gmember.GetType().InvokeMember("adspath",'GetProperty', $null, $gmember, $null)
            $MemberTable.name= "$name"
            $MemberTable.type= "$class"
            $MemberTable.adspath="$ads"
            $membertable.sid=$usersid.value
            

            if ($userSID -like "$domainsid*")
                {
                $MemberTable.domain = "$domain"
                }
            
            $MemberResult += $MemberTable
            }
            
         }
         $finalresult += $MemberResult 
    }
    $finalresult | select server,"local group",name,type,domain,sid
    }
}
## End Get-LocalGroupMembers
## Begin Get-LocalServerAdmins
Function Get-LocalServerAdmins {
# ============================================================================================== 
# NAME: Listing Administrators and PowerUsers on remote machines  
#  
# AUTHOR: Mohamed Garrana ,  
# DATE  : 09/04/2010 
#  
# COMMENT:  
# This script runs against an input file of computer names , connects to each computer and gets a list of the users in the  local Administrators  
#and powerusers Groups . the output can be a csv file which can be readable on excel with all the computers from the input file 
# ============================================================================================== 	
        param( 
    [Parameter(Mandatory=$true,valuefrompipeline=$true)] 
    [string]$strComputer) 
    begin {} 
    Process { 
        $adminlist ="" 
        #$powerlist ="" 
        $computer = [ADSI]("WinNT://" + $strComputer + ",computer") 
        $AdminGroup = $computer.psbase.children.find("Administrators") 
        #$powerGroup = $computer.psbase.children.find("Power Users") 
        $Adminmembers= $AdminGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        #$Powermembers= $PowerGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        foreach ($admin in $Adminmembers) { $adminlist = $adminlist + $admin + "," } 
        #foreach ($poweruser in $Powermembers) { $powerlist = $powerlist + $poweruser + "," } 
        $Computer = New-Object psobject 
        $computer | Add-Member noteproperty ComputerName $strComputer 
        $computer | Add-Member noteproperty Administrators $adminlist 
        #$computer | Add-Member noteproperty PowerUsers $powerlist 
        Write-Output $computer 
 
 
        } 
end {} 
} 
## End Get-LocalServerAdmins 
## Begin Get-LocalUser
Function Get-LocalUser{
	
<#
	.SYNOPSIS
		This script can be list all of local user account.
	
	.DESCRIPTION
		This script can be list all of local user account.
		The Function is using WMI to connect to the remote machine
	
	.PARAMETER ComputerName
		Specifies the computers on which the command . The default is the local computer.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	
	.EXAMPLE
		Get-LocalUser
		
		This example shows how to list all of local users on local computer.
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
#>
	
	PARAM
	(
		[Alias('cn')]
		[String[]]$ComputerName = $Env:COMPUTERNAME,
		
		[String]$AccountName,
		
		[System.Management.Automation.PsCredential]$Credential
	)
	
	$Splatting = @{
		Class = "Win32_UserAccount"
		Namespace = "root\cimv2"
		Filter = "LocalAccount='$True'"
	}
	
	#Credentials
	If ($PSBoundParameters['Credential']) { $Splatting.Credential = $Credential }
	
	Foreach ($Computer in $ComputerName)
	{
		TRY
		{
			Write-Verbose -Message "[PROCESS] ComputerName: $Computer"
			Get-WmiObject @Splatting -ComputerName $Computer | Select-Object -Property Name, FullName, Caption, Disabled, Status, Lockout, PasswordChangeable, PasswordExpires, PasswordRequired, SID, SIDType, AccountType, Domain, Description
		}
		CATCH
		{
			Write-Warning -Message "[PROCESS] Issue connecting to $Computer"
		}
	}
}
## End Get-LocalUser