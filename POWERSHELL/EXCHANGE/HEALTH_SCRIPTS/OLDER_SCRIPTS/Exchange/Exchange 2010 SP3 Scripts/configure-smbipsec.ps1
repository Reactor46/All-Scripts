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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUfNhLTkjhHJMa4x6Af+0u/L6y
# 3Q6gghhqMIIE2jCCA8KgAwIBAgITMwAAAR+XYwozuYPXKwAAAAABHzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM3
# WhcNMjAwMTEwMjEwNzM3WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046NDlCQy1FMzdBLTIzM0MxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCppklVnT29zi13dODY0ejMsdoe7n2iCvC6QdH5FJkRYfy+
# cXoHBmpDgDF/65Kt9GMmu/K8HKAzjKHeG18rgRXQagLwIIH5yCRbXGwOfuHIu1dC
# 26o/CT22+YlRvBJwH36WVjML8BLNDT3Fr+yhU4ZM7Hbegql4r5kSgsrrjyx5bJY5
# r2N0G7RDnbhRd79iqXbvDnvkatjB5xgluzfQEAPbJjXjmRb5685DEEZg1qFsQJer
# XuBA+ZVevuCX0DuDj8UmhHGC5Y32sulFTn283R6LU+8+AALtbHOOIHV7QHNYV8mN
# jxHuKLvE9tNEGIpbG2WF2yQkSGe3sRbGQmaILWeHAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUuPNVyPmK8/JJioMtQFlTUeF3IOgwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAmAYfr1fEosYv9VTf
# 0Msya6aFm0Id6Zq1O5jNy74ByTh7EEac/l/4e3DOyrczHS6zwvMKYzLtmifeGZvD
# 70qbbUfF+yjpzpyu00uuzZ1HNOpktp5/dJXkzz0NyVnEeFGOXLpNyZNIA9dKGDwN
# XbsEUukTX9lJFx5RcBhE8AOl22IHSgJ6NYf4DpATCjSJbC9IrKYGBchHobCLZHEt
# cLBjxXiWJRG2YY+LBAVW95gwNdPmLCKrob7SdNLK1VnM35Q2VgNF7YfDc5nw4E7C
# 4VaZvlyuDET6fYycIVPx5GsLhx3it4a+WKcBwarK7inH9skUArxMZrpWmjuQ/o4b
# GprEnjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# vDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUMOi4K2PKFLwOcCuOqQzC7LXbJ5sw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBADggRPo6/PvVPnw9NW9M+c3QU4t+PZNf1PisppcUmUeB
# JOoHjE+EoGRij3VyVT5xGfTKSudT4T9BZHm4y3rWFSS33JukDbmLSKTRSuet7Bpg
# YYZQIMhtlFksjKTT2qgerPmiKOV2z2AE3BOU7AOSxeHE9hSzbeoSr7tPk7c/bL+M
# Gwt8mwuIbhJLiJAGnvI3KdJS2txcpWFrawCj1Es2gMOCaooqCtllzxmLtSBNWmWG
# 5jOg1dg2am2k7cuEjqtTJUcplptRH/iHd1QRgewmfykgjBPpuv61/8YCKkyv92xy
# 6wg7ulOeRXKHP6VX5TZdpIEaJy3d77jOSs/XG2PTl32hggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# H5djCjO5g9crAAAAAAEfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NThaMCMGCSqGSIb3DQEJ
# BDEWBBSmJwkem8z+yyXgy4kIbVIPosJ5OTANBgkqhkiG9w0BAQUFAASCAQCJYnGg
# VOYpQ1LzAik3JBlk8SP7qXT9AtO0DYuDJ/EE5Gn1vhvvJ03oDF/YFZQ6MxN+sPDw
# G8r2bpOFOeyzu6fDE4XfsuIdIcdrixjHAsZV07on9AMp75QNDH3I9HyOBsao8FAB
# cfgl2g9gyPZ99sWVD4RGKat/j+Par226+JvmWmSTT8YhBDkVwasCY9HIkx6osWIn
# n5wex2hLOGbElknWwY/UaohaVpUyOTId1APBPi/9XI3R77LF4IatTJRyK1lNO6Pr
# Afk17rZ5xHvjlBFv+3EWHP/7txNh/awbrhIkYx3ybLyeEIQnmqOb3JZ+AgnAxg/J
# rlPIA4B+A20CAItj
# SIG # End signature block
