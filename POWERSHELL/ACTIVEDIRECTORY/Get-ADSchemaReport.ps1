<#-----------------------------------------------------------------------------
How to find AD schema update history using PowerShell
Ashley McGlone, Microsoft Premier Field Engineer
http://blogs.technet.com/b/ashleymcglone
December, 2011

This script reports on schema update and version history for Active Directory.
It requires the ActiveDirectory module to run.
It makes no changes to the environment.

UPDATED:
2013-03-12  Added Windows Server 2012, Exchange 2010 SP3 & 2013, Lync 2013
2013-09-19  Added Windows Server 2012 R2
            Added OID for schema attributes
            Added schema.csv output
            Sorted output
2014-06-12  Added Exchange 2013 SP1 as documented in the comments on the Scripting Guy blog
            Added SCCM as documented in the comments on the Scripting Guy blog
            http://blogs.technet.com/b/heyscriptingguy/archive/2012/01/05/how-to-find-active-directory-schema-update-history-by-using-powershell.aspx
2014-06-24  Added Exchange 2013 cumulative update versions as per feedback
            http://supertekboy.com/2014/05/01/check-exchange-schema-objects-adsi-edit/
            http://www.bhargavs.com/index.php/2009/11/20/verify-exchange-server-schema-version/

References for schema values:
http://support.microsoft.com/kb/556086?wa=wsignin1.0
http://social.technet.microsoft.com/wiki/contents/articles/2772.exchange-schema-versions-common-questions-answers.aspx
http://technet.microsoft.com/en-us/library/gg412822.aspx

LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
 
This posting is provided "AS IS" with no warranties, and confers no rights. Use
of included script samples are subject to the terms specified
at http://www.microsoft.com/info/cpyright.htm.
-----------------------------------------------------------------------------#>

Import-Module ActiveDirectory 

$schema = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) `
    -SearchScope OneLevel -Filter * -Property objectClass, name, whenChanged,`
    whenCreated, attributeID | Select-Object objectClass, attributeID, name,`
    whenCreated, whenChanged, `
    @{name="event";expression={($_.whenCreated).Date.ToString("yyyy-MM-dd")}} |
    Sort-Object event, objectClass, name

"`nDetails of schema objects created by date:"
$schema | Format-Table objectClass, attributeID, name, whenCreated, whenChanged `
    -GroupBy event -AutoSize

"`nCount of schema objects created by date:"
$schema | Group-Object event | Format-Table Count, Name, Group -AutoSize

$schema | Export-CSV .\schema.csv -NoTypeInformation
"`nSchema CSV output here: .\schema.csv"

#------------------------------------------------------------------------------

"`nForest domain creation dates:"
Get-ADObject -SearchBase (Get-ADForest).PartitionsContainer `
    -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3))" `
    -Property dnsRoot, nETBIOSName, whenCreated |
  Sort-Object whenCreated |
  Format-Table dnsRoot, nETBIOSName, whenCreated -AutoSize

#------------------------------------------------------------------------------

$SchemaVersions = @()

$SchemaHashAD = @{
    13="Windows 2000 Server";
    30="Windows Server 2003 RTM";
    31="Windows Server 2003 R2";
    44="Windows Server 2008 RTM";
    47="Windows Server 2008 R2";
    56="Windows Server 2012 RTM";
    69="Windows Server 2012 R2"
    }

$SchemaPartition = (Get-ADRootDSE).NamingContexts | Where-Object {$_ -like "*Schema*"}
$SchemaVersionAD = (Get-ADObject $SchemaPartition -Property objectVersion).objectVersion
$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"AD"}}, `
    @{name="Schema";expression={$SchemaVersionAD}}, `
    @{name="Version";expression={$SchemaHashAD.Item($SchemaVersionAD)}}

#------------------------------------------------------------------------------

$SchemaHashExchange = @{
    4397="Exchange Server 2000 RTM";
    4406="Exchange Server 2000 SP3";
    6870="Exchange Server 2003 RTM or SP1 or SP2";
    6936="Exchange Server 2003 SP3";
    10637="Exchange Server 2007 RTM";
    11116="Exchange 2007 SP1";
    14622="Exchange 2007 SP2 or Exchange 2010 RTM";
    14625="Exchange 2007 SP3";
    14726="Exchange 2010 SP1";
    14732="Exchange 2010 SP2";
    14734="Exchange 2010 SP3";
    15137="Exchange 2013 RTM";
    15254="Exchange 2013 CU1";
    15281="Exchange 2013 CU2";
    15283="Exchange 2013 CU3";
    15292="Exchange 2013 SP1"
    }

$SchemaPathExchange = "CN=ms-Exch-Schema-Version-Pt,$SchemaPartition"
If (Test-Path "AD:$SchemaPathExchange") {
    $SchemaVersionExchange = (Get-ADObject $SchemaPathExchange -Property rangeUpper).rangeUpper
} Else {
    $SchemaVersionExchange = 0
}

$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"Exchange"}}, `
    @{name="Schema";expression={$SchemaVersionExchange}}, `
    @{name="Version";expression={$SchemaHashExchange.Item($SchemaVersionExchange)}}

#------------------------------------------------------------------------------

$SchemaHashLync = @{
    1006="LCS 2005";
    1007="OCS 2007 R1";
    1008="OCS 2007 R2";
    1100="Lync Server 2010";
    1150="Lync Server 2013"
    }

$SchemaPathLync = "CN=ms-RTC-SIP-SchemaVersion,$SchemaPartition"
If (Test-Path "AD:$SchemaPathLync") {
    $SchemaVersionLync = (Get-ADObject $SchemaPathLync -Property rangeUpper).rangeUpper
} Else {
    $SchemaVersionLync = 0
}

$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"Lync"}}, `
    @{name="Schema";expression={$SchemaVersionLync}}, `
    @{name="Version";expression={$SchemaHashLync.Item($SchemaVersionLync)}}

#------------------------------------------------------------------------------

$SchemaHashSCCM = @{ 
    "4.00.5135.0000"="SCCM 2007 Beta 1";
    "4.00.5931.0000"="SCCM 2007 RTM";
    "4.00.6221.1000"="SCCM 2007 SP1/R2";
    "4.00.6221.1193"="SCCM 2007 SP1 (KB977203)";
    "4.00.6487.2000"="SCCM 2007 SP2";
    "4.00.6487.2111"="SCCM 2007 SP2 (KB977203)";
    "4.00.6487.2157"="SCCM 2007 R3";
    "4.00.6487.2207"="SCCM 2007 SP2 (KB2750782)";
    "5.00.7561.0000"="SCCM 2012 Beta 2";
    "5.00.7678.0000"="SCCM 2012 RC1";
    "5.00.7703.0000"="SCCM 2012 RC2";
    "5.00.7711.0000"="SCCM 2012 RTM";
    "5.00.7711.0200"="SCCM 2012 CU1";
    "5.00.7711.0301"="SCCM 2012 CU2";
    "5.00.7782.1000"="SCCM 2012 SP1 Beta";
    "5.00.7804.1000"="SCCM 2012 SP1";
    "5.00.7804.1202"="SCCM 2012 SP1 CU1";
    "5.00.7804.1300"="SCCM 2012 SP1 CU2";
    "5.00.7804.1400"="SCCM 2012 SP1 CU3";
    "5.00.7804.1500"="SCCM 2012 SP1 CU4";
    "5.00.7958.1000"="SCCM 2012 R2" 
    }

$SchemaPathSCCM = "CN=System Management," + (Get-ADDomain).SystemsContainer
if (Test-Path "AD:$SchemaPathSCCM") {
  $SCCMData = Get-ADObject -SearchBase ("CN=System Management," + (Get-ADDomain).SystemsContainer) -LDAPFilter "(&(objectClass=mSSMSManagementPoint))" -Property mSSMSCapabilities,mSSMSMPName
  if ($sccmdata -isnot [system.Array]) {
    $SCCMxml = [XML]$SCCMdata.mSSMSCapabilities
    $schemaVersionSCCM = $SCCMxml.ClientOperationalSettings.Version
  } else {
    $schemaVersionSCCMList = Foreach($SCCMINstance in 0..($SCCMData.count -1)) {
      $SCCMxml = [XML]$SCCMdata[$SCCMInstance].mSSMSCapabilities
      $SCCMxml.ClientOperationalSettings.Version
      }
    $SchemaVersionSCCM = $schemaVersionSCCMList|Sort-Object -Descending|Select-Object -First 1
  }
} Else {
  $schemaVersionSCCM = 0
}

$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"SCCM"}}, `
    @{name="Schema";expression={$schemaVersionSCCM}}, `
    @{name="Version";expression={$SchemaHashSCCM.Item($schemaVersionSCCM)}}

#------------------------------------------------------------------------------

"`nKnown current schema version of products:"
$SchemaVersions | Format-Table * -AutoSize

#---------------------------------------------------------------------------sdg
