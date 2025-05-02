# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

# Synopsis: This script is to be used to help add the necessary IPsec
#     configuration to protect SMB (File Share) communication.
#

# Define the script parameters
param([string] $Context                     = 'Static',
      [string] $Store                       = 'Local',
      [string] $PolicyName                  = 'SMB Security Policy (Exchange 2007)',
      [string] $OutputScriptFile            = '.\NetSH_Script_SMB.txt',
      [switch] $AddSMBServerFilterList      = $true,
      [switch] $AddSMBClientFilterList      = $true,
      [switch] $AddSMBServerFilterAction    = $true,
      [switch] $SMBServerFilterActionInPass = $true,
      [switch] $SMBServerFilterActionSoft   = $false,
      [switch] $AddSMBClientFilterAction    = $true,
      [switch] $SMBClientFilterActionInPass = $true,
      [switch] $SMBClientFilterActionSoft   = $true,
      [switch] $AddSMBServerRule            = $true,
      [switch] $AddSMBClientRule            = $true,
      [switch] $CreateSMBPolicy             = $true,
      [switch] $AssignSMBPolicy             = $false,
      [switch] $ViewNETSHScript             = $true,
      [switch] $ExecuteNETSHScript          = $false)


# PUSHD IPsec context command
$PUSHD_IPsec_Context = "pushd ipsec $Context"


# Set IPsec store command
$SET_IPsec_Store = "set store $Store"


# IPsec FilterList Name for SMB Server communications
$IPsec_SMB_Server_FilterList_Name = 'All SMB Traffic - Server'

# Add IPsec FilterList for SMB Server communications command
$ADD_IPsec_SMB_Server_FilterList = "add filterlist name=`"$IPsec_SMB_Server_FilterList_Name`" description=`"Matches all SMB packets for incoming SMB requests`""

# Add IPsec Filters for SMB Server communications commands
$ADD_IPsec_SMB_Server_Filters = @("add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - TCP 137`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=TCP srcport=0 dstport=137",
                                  "add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - UDP 137`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=UDP srcport=0 dstport=137",
                                  "add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - UDP 138`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=UDP srcport=0 dstport=138",
                                  "add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - TCP 139`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=TCP srcport=0 dstport=139",
                                  "add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - TCP 445`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=TCP srcport=0 dstport=445",
                                  "add filter filterlist=`"$IPsec_SMB_Server_FilterList_Name`" description=`"SMB Traffic - Incoming - UDP 445`" mirrored=yes srcaddr=ANY srcmask=0.0.0.0 dstaddr=ME dstmask=255.255.255.255 protocol=UDP srcport=0 dstport=445")


# IPsec FilterList Name for SMB Client communications
$IPsec_SMB_Client_FilterList_Name = 'All SMB Traffic - Client'

# Add IPsec FilterList for SMB Client communications command
$ADD_IPsec_SMB_Client_FilterList = "add filterlist name=`"$IPsec_SMB_Client_FilterList_Name`" description=`"Matches all SMB packets for outgoing SMB requests`""

# Add IPsec Filters for SMB Server communications commands
$ADD_IPsec_SMB_Client_Filters = @("add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - TCP 137`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=TCP srcport=0 dstport=137",
                                  "add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - UDP 137`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=UDP srcport=0 dstport=137",
                                  "add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - UDP 138`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=UDP srcport=0 dstport=138",
                                  "add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - TCP 139`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=TCP srcport=0 dstport=139",
                                  "add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - TCP 445`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=TCP srcport=0 dstport=445",
                                  "add filter filterlist=`"$IPsec_SMB_Client_FilterList_Name`" description=`"SMB Traffic - Outgoing - UDP 445`" mirrored=yes srcaddr=ME srcmask=255.255.255.255 dstaddr=ANY dstmask=0.0.0.0 protocol=UDP srcport=0 dstport=445")


# Add IPsec policy for SMB command
$ADD_IPsec_SMB_Policy = "add policy `"$PolicyName`" description=`"IPsec Security Policy to secure both SMB Server and SMB client communications.  The default filter actions are Server Require Security and Client Request Security.`" mmpfs=no activatedefaultrule=no mmsec=`"3DES-SHA1-2 3DES-SHA1-3`""


# Add SMB Server Filter Action
$IPsec_SMB_Server_FilterAction_Name = "SMB Server Filter Action"

# Check if unsecured communications are accepted
if ($SMBServerFilterActionInPass)
{
  # Unsecure communications accepted, but respond using IPsec
  $IPsec_SMB_Server_FilterAction_InPass = "yes"
}
else
{
  # Unsecure communications not accepted
  $IPsec_SMB_Server_FilterAction_InPass = "no"
}

# Check if Unsecure communications are allowed
if ($SMBServerFilterActionSoft)
{
  # Unsecured communications allowed
  $IPsec_SMB_Server_FilterAction_Soft = "yes"
}
else
{
  # Unsecured communications not allowed
  $IPsec_SMB_Server_FilterAction_Soft = "no"
}

$ADD_IPsec_SMB_Server_FilterAction ="add filteraction name=`"$IPsec_SMB_Server_FilterAction_Name`" description=`"By Default, Require Security`" qmpfs=no inpass=$IPsec_SMB_Server_FilterAction_InPass soft=$IPsec_SMB_Server_FilterAction_Soft action=negotiate qmsec=`"ESP[3DES,SHA1]`""


# Add SMB Client Filter Action
$IPsec_SMB_Client_FilterAction_Name = "SMB Client Filter Action"

# Check if unsecured communications are accepted
if ($SMBClientFilterActionInPass)
{
  # Unsecure communications accepted, but respond using IPsec
  $IPsec_SMB_Client_FilterAction_InPass = "yes"
}
else
{
  # Unsecure communications not accepted
  $IPsec_SMB_Client_FilterAction_InPass = "no"
}

# Check if Unsecure communications are allowed
if ($SMBClientFilterActionSoft)
{
  # Unsecured communications allowed
  $IPsec_SMB_Client_FilterAction_Soft = "yes"
}
else
{
  # Unsecured communications not allowed
  $IPsec_SMB_Client_FilterAction_Soft = "no"
}

$ADD_IPsec_SMB_Client_FilterAction ="add filteraction name=`"$IPsec_SMB_Client_FilterAction_Name`" description=`"By Default, Request Security`" qmpfs=no inpass=$IPsec_SMB_Client_FilterAction_InPass soft=$IPsec_SMB_Client_FilterAction_Soft action=negotiate qmsec=`"ESP[3DES,SHA1]`""


# Add IPsec SMB Server rule to Policy command
$IPsec_SMB_Server_Rule_Name = 'SMB Server Rule'
$ADD_IPsec_SMB_Server_Rule = "add rule name=`"$IPsec_SMB_Server_Rule_Name`" policy=`"$PolicyName`" filterlist=`"$IPsec_SMB_Server_FilterList_Name`" filteraction=`"$IPsec_SMB_Server_FilterAction_Name`" conntype=all activate=yes description=`"By Default, Require Security for all Incoming SMB Traffic`" kerberos=yes"


# Add IPsec SMB Client rule to Policy command
$IPsec_SMB_Client_Rule_Name = 'SMB Client Rule'
$ADD_IPsec_SMB_Client_Rule = "add rule name=`"$IPsec_SMB_Client_Rule_Name`" policy=`"$PolicyName`" filterlist=`"$IPsec_SMB_Client_FilterList_Name`" filteraction=`"$IPsec_SMB_Client_FilterAction_Name`" conntype=all activate=yes description=`"By Default, Request Security for all Outgoing SMB Traffic`" kerberos=yes"


# Assign IPsec policy command
$SET_IPsec_Policy_Assign = "set policy `"$PolicyName`" assign=yes"


# NETSH Script Suffix
$NETSH_Script_Suffix='popd`r`nexit'


# This function validates the scripts parameters
function ValidateParams
{
  $validInputs = $true
  $errorString =  '`n`n################################################################################`n'
  $errorString += '# There were errors validating the script parameters!                          #`n'
  $errorString += '################################################################################`n'

  # Validate IPsec context
  if (!($Context -imatch "static|dynamic"))
  {
    $validInputs = $false
    $errorString += "`nERROR: The `"Context`" parameter must be `"static`" or `"dynamic`".`nSpecified Value: `"$Context`"`n"
  }

  # Validate IPsec Store
  if (!($Store -imatch "local|persistent|domain"))
  {
    $validInputs = $false
    $errorString += "`nERROR: The `"Store`" parameter must be `"local`",`"domain`", or `"persistent`".`nSpecified Value: `"$Store`"`n"
  }

  if (!$validInputs)
  {
    Write-Warning "$errorString`n`n"
  }

  return $validInputs
}


function WriteOutputFile
{
  param([string] $outputString,
        [bool]   $appendString   = $true,
        [string] $outputEncoding = "ASCII")

  if ($appendString)
  {
    $outputString | out-file $OutputScriptFile -Encoding $outputEncoding -Append
  }
  else
  {
    $outputString | out-file $OutputScriptFile -Encoding $outputEncoding
  }
}


function Usage()
{
@"

********************************************************************************

DISCLAIMER:

Careful consideration should be taken when deploying IPsec.  Testing IPsec
changes in a non-production environment is strongly recommended before deploying
the changes in a production environment.

It should be understood that the generated NETSH script may not contain the
correct IPsec policy, rules, filter lists, and/or filters for your organization.
Therefore, if additional customizations are required, simply modify the
generated NETSH script file and then run it manually or take the information
provided by the NETSH script file and manually create IPsec policies, rules,
filter lists, and filters that best apply to your organization.

********************************************************************************

SUMMARY:

Exchange 2007 uses file shares (Server Message Block - SMB) to transmit data
from one server to another.  Since some of this data may be "private" in nature,
it is necessary to secure the data while being transmitted across the network
between Exchange servers.  Currently, the recommended way to secure SMB
communications is by using IPsec.  Here are the file shares for Exchange 2007:

FILE SHARE NAME`t`tROLES`t`t`tDESCRIPTION

Address`t`t`tMailbox`t`t`tThis File Share contains the
`t`t`t`t`t`tproxy generation DLLs for the
`t`t`t`t`t`tlocal system.  The "Microsoft
`t`t`t`t`t`tExchange System Attendant"
`t`t`t`t`t`tservice on Exchange 2000, 2003,
`t`t`t`t`t`tand 2007 servers accesses this
`t`t`t`t`t`tFile Share on other Exchange
`t`t`t`t`t`tservers to check if they have a
`t`t`t`t`t`tnewer version of the proxy
`t`t`t`t`t`tgeneration DLLs.  If they do,
`t`t`t`t`t`tthe newer versions are copied
`t`t`t`t`t`tover.  There is no personal
`t`t`t`t`t`tdata stored in this file share.

ExchangeOAB`t`tMailbox`t`t`tThis File Share is utilized by
`t`t`t`t`t`tthe "Microsoft Exchange File
`t`t`t`t`t`tDistribution" service on the
`t`t`t`t`t`tExchange 2007 Client Access
`t`t`t`t`t`tservers to replicate the
`t`t`t`t`t`tExchange Offline Address Book(s)
`t`t`t`t`t`tfrom the Exchange 2007 Mailbox
`t`t`t`t`t`tserver(s).

ExchangeUM`t`tUnified Messaging`tThis File Share is utilized by
`t`t`t`t`t`tthe "Microsoft Exchange File
`t`t`t`t`t`tDistribution" service on
`t`t`t`t`t`tExchange 2007 Unified Messaging
`t`t`t`t`t`tservers to replicate the custom
`t`t`t`t`t`tUM prompts.  There is no
`t`t`t`t`t`tpersonal data stored in this
`t`t`t`t`t`tfile share.

<GUID>`t`t`tMailbox with CCR`tThis File Share is utilized by
`t`t`t`t`t`tthe "Microsoft Exchange
`t`t`t`t`t`tReplication Service" on Exchange
`t`t`t`t`t`t2007 Mailbox servers to copy the
`t`t`t`t`t`tStorage Group transaction logs
`t`t`t`t`t`tfrom the active node in the CCR
`t`t`t`t`t`tCluster Pair to the passive
`t`t`t`t`t`tnode.  CCR stands for Continuous
`t`t`t`t`t`tCluster Replication.

********************************************************************************

USAGE:

configure-SMBIPsec.msh  [-Context "static|dynamic"]
                        [-Store   "local|domain|persistent"]
                        [-PolicyName <System.String>]
                        [-OutputScriptFile <System.String>]
                        [-AddSMBServerFilterList[:<System.Boolean>]]
                        [-AddSMBClientFilterList[:<System.Boolean>]]
                        [-AddSMBServerFilterAction[:<System.Boolean>]]
                        [-SMBServerFilterActionInPass[:<System.Boolean>]]
                        [-SMBServerFilterActionSoft[:<System.Boolean>]]
                        [-AddSMBClientFilterAction[:<System.Boolean>]]
                        [-SMBClientFilterActionInPass[:<System.Boolean>]]
                        [-SMBClientFilterActionSoft[:<System.Boolean>]]
                        [-AddSMBServerRule[:<System.Boolean>]]
                        [-AddSMBClientRule[:<System.Boolean>]]
                        [-CreateSMBPolicy[:<System.Boolean>]]
                        [-AssignSMBPolicy[:<System.Boolean>]]
                        [-ViewNETSHScript[:<System.Boolean>]]
                        [-ExecuteNETSHScript[:<System.Boolean>]]

-Context`t`tSpecifies whether to use the 'Static' or 'Dynamic' IPsec
`t`t`tcontext.  'Static' allows you to create, modify, and
`t`t`tassign IPsec polices without affecting the
`t`t`tconfiguration of the active IPsec policy.  'Dynamic', on
`t`t`tthe other hand, affects the configuration of the active
`t`t`tIPsec policy.  Default value is '$Context'.

-Store`t`t`tSpecifies whether to use the 'Local', 'Domain', or
`t`t`t'Persistent' IPsec store.  'Local' refers to the IPsec
`t`t`tstore on the local computer.  'Domain' refers to the
`t`t`tIPsec store for the domain.  'Persistent' refers to the
`t`t`tIPsec store on the local computer that contains policies
`t`t`tto secure the computer on start up, before the local
`t`t`tpolicy or domain-based policy is applied.  Default value
`t`t`tis '$Store'.

-PolicyName`t`tSpecifies the name of the IPsec policy that is
`t`t`tto be created, assigned, or modified by adding the
`t`t`tappropriate rules.  If you have an existing IPsec policy
`t`t`tthat you would like to add the SMB Server and Client
`t`t`trules to, then you would specify that policy name here.
`t`t`tDefault value is '$PolicyName'.

-OutputScriptFile`tSpecifies the name of the output script
`t`t`tfile that will contain the appropriate NETSH.exe
`t`t`tcommands to make the specified IPsec modifications.
`t`t`tDefault value is '$OutputScriptFile'.

-AddSMBServerFilterList`tSpecifies that the IPsec FilterList, and corresponding
`t`t`tFilters, which matches all incoming SMB requests is to
`t`t`tbe added to the specified IPsec store.  The name of the
`t`t`tFilterList to be added is '$IPsec_SMB_Server_FilterList_Name'.
`t`t`tDefault value is '$AddSMBServerFilterList'.

-AddSMBClientFilterList`tSpecifies that the IPsec FilterList, and corresponding
`t`t`tFilters, which matches all outgoing SMB requests is to
`t`t`tbe added to the specified IPsec store.  The name of the
`t`t`tFilterList to be added is '$IPsec_SMB_Client_FilterList_Name'.
`t`t`tDefault value is '$AddSMBClientFilterList'.

-AddSMBServerFilterAction`tSpecifies that the IPsec Filter Action named
`t`t`t'$IPsec_SMB_Server_FilterAction_Name' is to be added.
`t`t`tThis Filter Action will either "Request" or "Require"
`t`t`tthe client to use IPsec depending on the values
`t`t`tspecified for the "-SMBServerFilterActionInPass" and
`t`t`t"-SMBServerFilterActionSoft" parameters.  This Filter
`t`t`tAction will use "3DES" for ESP Confidentiality and
`t`t`t"SHA1" for ESP Integrity.

-SMBServerFilterActionInPass`tSpecifies if the setting "Accept unsecured
`t`t`tcommunication, but always respond using IPsec" is to be
`t`t`tenabled for the SMB Server IPsec Filter Action.
`t`t`tDefault value is '$SMBServerFilterActionInPass'.

-SMBServerFilterActionSoft`tSpecifies if the setting "Allow unsecured
`t`t`tcommunications with non-IPsec-aware computers" is to be
`t`t`tenabled for the SMB Server IPsec Filter Action.
`t`t`tDefault value is '$SMBServerFilterActionSoft'.

-AddSMBClientFilterAction`tSpecifies that the IPsec Filter Action named
`t`t`t'$IPsec_SMB_Client_FilterAction_Name' is to be added.
`t`t`tThis Filter Action will either "Request" or "Require"
`t`t`tthe client to use IPsec depending on the values
`t`t`tspecified for the "-SMBClientFilterActionInPass" and
`t`t`t"-SMBClientFilterActionSoft" parameters.  This Filter
`t`t`tAction will use "3DES" for ESP Confidentiality and
`t`t`t"SHA1" for ESP Integrity.

-SMBClientFilterActionInPass`tSpecifies if the setting "Accept unsecured
`t`t`tcommunication, but always respond using IPsec" is to be
`t`t`tenabled for the SMB Client IPsec Filter Action.
`t`t`tDefault value is '$SMBClientFilterActionInPass'.

-SMBClientFilterActionSoft`tSpecifies if the setting "Allow unsecured
`t`t`tcommunications with non-IPsec-aware computers" is to be
`t`t`tenabled for the SMB Client IPsec Filter Action.
`t`t`tDefault value is '$SMBClientFilterActionSoft'.

-AddSMBServerRule`tSpecifies that the IPsec Rule named '$IPsec_SMB_Server_Rule_Name'
`t`t`tis to be added to the specified IPsec policy.  This Rule
`t`t`twill contain the '$IPsec_SMB_Server_FilterList_Name' FilterList
`t`t`tand will apply to the Filter List the Filter Action named
`t`t`t'$IPsec_SMB_Server_FilterAction_Name'.
`t`t`tDefault value is '$AddSMBServerRule'.

-AddSMBClientRule`tSpecifies that the IPsec Rule named '$IPsec_SMB_Client_Rule_Name'
`t`t`tis to be added to the specified IPsec policy.  This Rule
`t`t`twill contain the '$IPsec_SMB_Client_FilterList_Name' FilterList
`t`t`tand will apply to the Filter LIst the Filter Action named
`t`t`t'$IPsec_SMB_Client_FilterAction_Name'.
`t`t`tDefault value is '$AddSMBClientRule'.

-CreateSMBPolicy`tSpecifies that the IPsec policy specified in the
`t`t`t'-PolicyName' parameter is to be created.
`t`t`tDefault value is '$CreateSMBPolicy'.

-AssignSMBPolicy`tSpecifies that the IPsec policy specified in the
`t`t`t'-PolicyName' parameter is to be assigned.  Be aware
`t`t`tthat only one IPsec policy can be assigned to a computer
`t`t`tat a time.  Also, if you have specified the value
`t`t`t'Domain' for the '-Store' parameter, this command will
`t`t`thave not affect.  Default value is '$AssignSMBPolicy'.

-ViewNETSHScript`tSpecifies that the output script file is to be viewed
`t`t`tusing NOTEPAD.exe when completed.  Default value '$ViewNETSHScript'.

-ExecuteNETSHScript`tSpecifies that the output script file is to be executed
`t`t`tby NETSH.exe.  If the script is executed, then the
`t`t`t"ViewNETSHScript" parameter will be set to `$false.
`t`t`tDefault value is '$ExecuteNETSHScript'.

********************************************************************************

EXAMPLES:

1.)  View the NETSH commands to create local IPsec policy for SMB:
       .\configure-SMBIPsec.ps1

2.)  Import SMB IPsec settings to the Local store:
       .\configure-SMBIPsec.ps1 -AssignSMBPolicy -ExecuteNETSHScript

3.)  Import SMB IPsec settings to an existing Domain IPsec Policy:
       .\configure-SMBIPsec.ps1 -Store "Domain" -PolicyName "Contoso IPsec Policy" -CreateSMBPolicy:`$false -ExecuteNETSHScript

4.)  Import SMB IPsec settings for outgoing SMB requests:
       .\configure-SMBIPsec.ps1 -AssignSMBPolicy -AddSMBServerFilterList:`$false -AddSMBServerFilterAction:`$false -AddSMBServerRule:`$false -ExecuteNETSHScript

********************************************************************************

ADDITIONAL INFORMATION:

If you look at the generated NETSH script file, you will notice that there are
six ports covered by each of the FilterLists.  The reason for this is because
SMB communication occurs over ports TCP/UDP 445 as well as ports TCP/UDP 137,
UDP 138, and TCP 139 when "NetBIOS over TCP/IP" is enabled.  The only way to
force SMB communications to always occurs over ports TCP/UDP 445 is to disable
"NetBIOS over TCP/IP".

To find out more information about disabling "NetBIOS over TCP/IP" for the
direct hosting of SMB over TCP/IP, please refer to the references section below.

                                   ++++++++++

It should be noted that the default behavior for the SMB Client Filter Action is
"Request" and the default behavior for the SMB Server Filter Action is
"Require".  By default, both Filter Actions accept unsecured communications but
only the Client Filter Action allows for falling back to allow unsecure
communications.

This means that all incoming SMB requests will have to use IPsec to secure the
SMB communications.  Outgoing SMB requests will attempt to use IPsec to secure
the SMB communications, but if the remote computer does not support IPsec, the
the communications will fall back to being in the clear.

"NetBIOS over TCP/IP" is used by many applications and not just SMB
communications.  Great care should be taken when deploying these IPsec settings
to ensure that other applications and servers are not adversely affected.

                                   ++++++++++

NETSH.exe has the ability to execute commands on a remote server.  If you wanted
to modify the IPsec settings on a remote computer, you could simply run:
  NETSH.exe -r <REMOTE_COMPUTER_NAME> -f "$OutputScriptFile"

With some help of some of the Exchange 2007 commandlets, you can take this one
step further to deploy the IPsec settings to all Exchange servers:
  Get-ExchangeServer | foreach(`$_) {NETSH.exe -r `$_.Name -f "$OutputScriptFile"}

Or you could use 'Get-ClientAccessService', 'Get-MailboxServer',
'Get-TransportService', or 'Get-UmService' to deploy the IPsec settings to only
specific roles.

********************************************************************************

REFERENCES:

Deploying IPsec
http://technet2.microsoft.com/WindowsServer/en/Library/0bd06cf7-2ed6-46f1-bb55-2bf870273e151033.mspx

Server and Domain Isolation
http://www.microsoft.com/sdisolation

NETSH commands for Internet Protocol security
http://technet2.microsoft.com/WindowsServer/en/Library/c3ae0d03-f18f-40ac-ad33-c0d443d5ed901033.mspx

Overview of Server Message Block Signing (SMB)
http://support.microsoft.com/kb/887429

Direct hosting of SMB over TCP/IP
http://support.microsoft.com/kb/Q204279

Microsoft Windows Server 2003 TCP/IP Implementation Details
http://technet2.microsoft.com/WindowsServer/en/Library/823ca085-8b46-4870-a83e-8032637a87c81033.mspx

TCP/IP Fundamentals for Microsoft Windows : Chapter 11 - NetBIOS over TCP/IP
http://www.microsoft.com/technet/itsolutions/network/evaluate/technol/tcpipfund/tcpipfund_ch11.mspx

********************************************************************************

"@
}


####################################################################################################
# Script starts here
####################################################################################################

# Check for Usage Statement Request
if (($args.Count -gt 0) -and ($args[0] -imatch "-{1,2}[?h]"))
{
  # User wants the Usage Statement
  Usage
  return
}

# Validate the parameters
$ifValidParams = ValidateParams

if ($ifValidParams -eq $true)
{
  # Valid parameters

  # Add comment to output script file
  WriteOutputFile "# Execute this script by running 'NETSH.exe -f `"$OutputScriptFile`"'" $false

  # Specify the IPsec Context
  WriteOutputFile $PUSHD_IPsec_Context

  # Specify the IPsec Store
  WriteOutputFile $SET_IPsec_Store

  # Add the SMB Server FilterList
  if ($AddSMBServerFilterList)
  {
    # Create the FilterList
    WriteOutputFile $ADD_IPsec_SMB_Server_FilterList
    
    # Add the filters to the FilterList
    foreach ($filter in $ADD_IPsec_SMB_Server_Filters)
    {
      WriteOutputFile $filter
    }
  }

  # Add the SMB Client FilterList
  if ($AddSMBClientFilterList)
  {
    # Create the FilterList
    WriteOutputFile $ADD_IPsec_SMB_Client_FilterList

    # Add the filters to the FilterList
    foreach ($filter in $ADD_IPsec_SMB_Client_Filters)
    {
      WriteOutputFile $filter
    }
  }

  # Add the SMB Server FilterAction
  if ($AddSMBServerFilterAction)
  {
    # Create the FilterAction
    WriteOutputFile $ADD_IPsec_SMB_Server_FilterAction
  }

  # Add the SMB Client FilterAction
  if ($AddSMBClientFilterAction)
  {
    # Create the FilterAction
    WriteOutputFile $ADD_IPsec_SMB_Client_FilterAction
  }

  # Create the default Policy
  if ($CreateSMBPolicy)
  {
    WriteOutputFile $ADD_IPsec_SMB_Policy
  }

  # Add the SMB Server Rule to the Policy
  if ($AddSMBServerRule)
  {
    WriteOutputFile $ADD_IPsec_SMB_Server_Rule
  }

  # Add the SMB Client Rule to the Policy
  if ($AddSMBClientRule)
  {
    WriteOutputFile $ADD_IPsec_SMB_Client_Rule
  }

  # Assign the Policy
  if ($AssignSMBPolicy -and ($Store -ine "Domain"))
  {
    WriteOutputFile $SET_IPsec_Policy_Assign
  }

  # Append the NETSH Suffix
  WriteOutputFile $NETSH_Script_Suffix

  # Execute NETSH Script
  if ($ExecuteNETSHScript)
  {
    write-host "`nExecuting 'NETSH.exe -f `"$OutputScriptFile`"'"
    NETSH.exe -f "$OutputScriptFile"
    write-host "`n"
  }
  else
  {
    # View NETSH Script
    if ($ViewNETSHScript)
    {
      write-host "`nExecuting 'NOTEPAD.exe `"$OutputScriptFile`"'"
      NOTEPAD.exe "$OutputScriptFile"
      write-host "`n"
    }
  }
}

# SIG # Begin signature block
# MIIdrAYJKoZIhvcNAQcCoIIdnTCCHZkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUl9HutcIU1ja2b2MHiIJw0wO7
# fa6gghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBLIwggSuAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBxjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU98JsDqVmKSekXVU3bmA7lC3uxAIwZgYKKwYB
# BAGCNwIBDDFYMFagLoAsAGMAbwBuAGYAaQBnAHUAcgBlAC0AUwBNAEIASQBQAHMA
# ZQBjAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdl
# IDANBgkqhkiG9w0BAQEFAASCAQADFIp00EqYWPZjSA/AMed9GmuhqO18dVZef5o/
# 9X7jRnLHXipdWFrNuoslWiF4sMp1rtWLdn9BeRJoNhAreiWi/TCT7A5oB0IGkE81
# RLndfNDuBU8oHR6WyOACrEIFgcTbuYgSTBN3Ph/dMX2lddlgcLCHcRKWOigQYth3
# sGw8Ts59pjRt/vlK125SmGmdXpHbwwaKd8sIi4c1WU5yUTjpLUS30772itBmmsho
# F+rD3r/vWcbX6bcpL35Wn2ZuUjsSKv/twoPO3y/7B2lzHQaUrM90J3HQ3Cott5Ns
# OsO91AAQjBzDET3vjjeBHwd7b4lUxqvC8A9iAtzPeiGYWzFeoYICKDCCAiQGCSqG
# SIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQIT
# MwAAAJmqxYGfjKJ9igAAAAAAmTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsG
# CSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0NDM5WjAjBgkqhkiG
# 9w0BCQQxFgQUxmajFLxEIO9m2hudX1hws7ueP+4wDQYJKoZIhvcNAQEFBQAEggEA
# fxCt6byo05u2q5r23l3sAEswqPmo3ZyCT3aPKEzQuQC9YAV+v8V6S406yWJNhyno
# ttiD8i9QgzDvH+8pJngtRuCZ82vR9xHS8ynkCzDzzxlMaTPBlT8GyhsMCvtdYR0X
# uTR1svq3PzbnxP8Uoeqe33cIT2qL3IqiUfSkA82vrWjaqUF+i4ETUREndElRO7lw
# fwZlX0JsPmrbMyd5OmF7B2aYNPY9OPSgPARuMvweMfrwIEOT+iC5EWpqQb7ses68
# q3fJyjwkHLy+LYaoBOOqTL9YriZVkdpZ4Jd921reY2AdaOpVtkN/6rUB6RCamEtR
# 0/GBptAcmPaOrOSw5/xqoA==
# SIG # End signature block
