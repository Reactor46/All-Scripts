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

Or you could use 'Get-ClientAccessServer', 'Get-MailboxServer',
'Get-TransportServer', or 'Get-UmServer' to deploy the IPsec settings to only
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
# MIIacgYJKoZIhvcNAQcCoIIaYzCCGl8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfNhLTkjhHJMa4x6Af+0u/L6y
# 3Q6gghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggStMIIEqQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHGMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBQw6LgrY8oUvA5wK46pDMLstdsnmzBmBgorBgEEAYI3AgEMMVgw
# VqAugCwAYwBvAG4AZgBpAGcAdQByAGUALQBTAE0AQgBJAFAAcwBlAGMALgBwAHMA
# MaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3
# DQEBAQUABIIBAKdQUR9VLUT4umMQtsbilhUTb4uKxhi7uUJPMjerfoQJ8vFTH5P5
# SJyCo6w2xYuShDvW/VjqRXttaM/KU+78viqng2RW8Uswb8wPWdBmWtpB/zVtrMBH
# Ra1VIZLmTIhhztOm+s3e4yjTBCeAE3Zn5ETU9JmdB5ok6I6TGAstGbo6eXmdWR2G
# nxmiU7S/1ipCjfFUc2HDfWgvXnyioh/kAEuG3pn0PurHzDpMzt3Jh5U8IbmG1ca7
# +OSFdgBOKqIduj7yjxC/ZM/H9VBx/kVE0uB/BPWYqGeQexhV0bZ1fGquJpmM+i1X
# FprA0v9+RFoLvlmfd8TqKB2vG/Y+SgDRTYqhggIoMIICJAYJKoZIhvcNAQkGMYIC
# FTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAKzkySMGy
# yUjzAAAAAAArMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xMzAyMDUwNjM3MjFaMCMGCSqGSIb3DQEJBDEWBBSC
# tdR+QzgkEuQKrYIMTlvqo10brTANBgkqhkiG9w0BAQUFAASCAQA4qAqbFHBS+IkP
# E0eRLfT/Clo9K5reqxUc4X+vu4zY7iKJd0sJ4+OYgONgFmPb0gMA3baZ8t0+ERBe
# vEuPFxUyy/oGHi5X7n8dtXLRH/HDiRLBjfpUyU0Vrn7Wv2vNHTamSIUxjrneAo+I
# bY9LTf4LqnyY6V7MbYYXmDMttQ2enypva6yO6xtxcD5+9XzK85AjOIL0msFVyi2n
# 4ydDac7h4gguzg8dzwKPF3bwCREia97d31oNpli791n59z0ofdb2Lu2U+L+u24dW
# c0qYl1/9hkY9TjXOrBvcU+AHUq9Qio+5yvmkVX2Ykea8Cb8qcJ+Iu4UhoKHh+dXa
# +GWoKoSj
# SIG # End signature block
