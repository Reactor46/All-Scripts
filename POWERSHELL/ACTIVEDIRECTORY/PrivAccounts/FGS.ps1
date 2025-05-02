<#
Script Info

Author: Pierre Audonnet [MSFT]
Blog: http://blogs.technet.com/b/pie/
Download: https://gallery.technet.microsoft.com/Administrative-Groups-53df8e65

Disclaimer:
This sample script is not supported under any Microsoft standard support program or service. 
The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims 
all implied warranties including, without limitation, any implied warranties of merchantability 
or of fitness for a particular purpose. The entire risk arising out of the use or performance of 
the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
or anyone else involved in the creation, production, or delivery of the scripts be liable for any 
damages whatsoever (including, without limitation, damages for loss of business profits, business 
interruption, loss of business information, or other pecuniary loss) arising out of the use of or 
inability to use the sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages
#>
<#
.Synopsis
    This script collects group memberships for the default administrative groups as well as information about the current forest
    or a specified forest.
.DESCRIPTION
    This script collects the members of all default groups of each domain of the current forest. This can be changed and it is possible
    to specify another forest as long as it is trusted. For each members it will collect all the attributes useful for a security review.
    It also collects information about the forest such as the schema modification or critical attributes such as dSHeuristics.
.EXAMPLE
    Recommended parameters 
    .\FGS.ps1 -VerboseOutput $true -DebugOutput $true
.EXAMPLE
    Complete command line
    .\FGS.ps1 -VerboseOutput $true -DebugOutput $true -SkipSchema $false -SkipTrust $false -SkipExport $false
.EXAMPLE
    To target another forest than the one of the user
    .\FGS.ps1 -VerboseOutput $true -DebugOutput $true -Forest "fabrikam.com"
.EXAMPLE
    To add a group to the collection even if it is not an admin group
    .\FGS.ps1 -AddGroups @("CONTOSO\Group 1")
    
    To add several groups to the collection even if they are not admin groups
    .\FGS.ps1 -AddGroups @("CONTOSO\Group 1","FABRIKAM\Group 2")
 
    or
 
    .\FGS.ps1 -AddGroups "CONTOSO\Group 1","FABRIKAM\Group 2"
 
.INPUTS
    -VerboseOutput [$true|$false]
        If $true, it includes the verbose information into the transcript. This is a recommended parameter because it facilitate
        the troubleshouting.
    
    -DebugOutput [$true|$false]
        If $true, it includes the debug information into the transcript. This is a recommended parameter because it facilitate
        the troubleshooting.

    -Forest [string]
        If specified, the script will target the specified forest instead of the forest of the user running the script.

    -SkipSchema [$true|$false]
        For test purposes only. If $true the schema collection is not performed. This generate error during the HTML
        report generation
    
    -SkipTrust [$true|$false]
        For test purposes only. If $true the trust collection is not performed. 

    -WideTrickyUAC [$true|$false]
        Some environement have many accounts with a password never set or a password that never expires. If $false we do not look for those.
        This is $true by default
 
    -SkipExport [$true|$false]
        For test purposes only. If $true script does not export the collected data into XML files. This means it becomes impossible
        to generate the reports offsite
 
    -AddGroups [@array]
        If used, it must contain an array.
.OUTPUTS
   If the -SkipExport $false the script will generate XML files containing all the information which are used to generated reports
   even if we are not connected to the environment.
.NOTES
    Version Tracking
    9:36 AM 2014-01-30
    Version 1.1  
        - Add version tracking
        - Add CmdLet help and examples
        - Add type to the input parameters
        - Fix report issue with the tombstoneLifetime value
        - Change the logic behind the calculation of the account with a pasword that never changed since the creation
        - Change the issue structure use a generic title and an create an objects property
        - Change the thresholds warning and error limits for the groups
    9:21 AM 2014-02-03
    Version 1.2
        - Add the switch $WideTrickyUAC
        - Fix the text in the progress bar for tricky UAC collection
        - Fix the issue ht for wellknown security RID in a group
        - Add $_default_attributes but do not test against it yet
        - Fix the $_issues export
        - Add export of the $_ht_admincount variable
    11:38 AM 2014-02-24
    Version 1.3
        - Fix the collection when custom groups are defined with -AddGroups
        - Remove the -AddUsers parameter
        - Disable the debug output for the HTML schema report generatio
        - Review the LimitError and LimitWarning default value for the colored bar graph
    09:11 AM 2014-03-12
    Version 1.4
        - Add the possibility to scan another forest as long as it is trsuted with the -Forest
    06:17 AM 2014-11-21
    Version 1.5
        - Correct a bug when a builtin group collect hit a DC from the wrong domain
    02:31 PM 2014-11-22
    Version 1.6
        - Correct a bug when the forest is not specified and interpreted as "" instead of $null
        - Correct a bug when the computer's name of the collection was displayed instead of the user's name
        - Correct a bug when the Guest group is displayed twice
    11:02 AM 2014-11-23
    Version 1.7
        - Catch irregularities in the $Forest input parameter and exist if it cannot access to the specified forest
    05:36 PM 2014-11-23
    Version 1.8
        - Add the Read-only Domain Controllers group to the collection
        - Add error management to the trust collection
        - Remove the delta column for the report section Accounts with a password that never changed
        - Add a whenCreated column the report section  Accounts with a password that never changed
#>

param(
    [bool]   $VerboseOutput = $false,  # Determine if the Write-Verbose will be displayed on the screen
    [bool]   $DebugOutput   = $false,  # Determine if the Write-Debug will be displayed on the screen
    [string] $Forest        = $null,  # specify the forest name if no the default one
    [bool]   $SkipSchema    = $false, # If $True, we skip the verification of the defaultSecurityDescriptor of the schema
    [bool]   $SkipTrust     = $false, # If $True, we skip the scan of all external trusts
    [bool]   $WideTrickyUAC = $true,  # If $False, we do not look for all the accounts with a password never set nor password never expires
    [bool]   $SkipExport    = $true, # If $True, we skip the generation of the XML files
    #$SkipHTML       = $false, # If $True, we skip the generation of the final HTML report # Not implemented
    [array]  $AddUsers      = $null,  # Should be an array of DOMAIN\SAMACCOUNTNAME, add the users the the $_ht_sid_objects
    [array]  $AddGroups     = $null   # Should be an array of DOMAIN\GROUPNAME, add the groups to the group collection with Admins=1 for the group
)
$_ScriptVersion   = "1.8"
$_ScriptStartDate = [System.DateTime]::Now
If ( $VerboseOutput -eq $true ) { $VerbosePreference = "Continue" } Else { $VerbosePreference = "SilentlyContinue"} 
If ( $DebugOutput -eq $true )   { $DebugPreference = "Continue"   } Else { $DebugPreference = "SilentlyContinue"  }
$_timestamp = (Get-Date).toString(‘yyyyMMddhhmm’)
$_collection_export_prefix = "$($_timestamp)_Clixml_"
Start-Transcript -Path "$($_timestamp)_transcript.txt"
# Init hashtables
$script:_ht_sid_domain  = @{} # Stores all domains and their related info (this includes memerships)
$script:_ht_sid_objects = @{} # Stores all members of the groups we are analyzing (this includes their certain of their attributes)
$_ht_admincount         = @{} # Stores the orphan protected objects
 
$script:_issues = @() # Stores all issues detected
 
# Display who and from where the collection is performed
$_collection_machine_name   = (Get-WmiObject Win32_ComputerSystem).Name
$_collection_machine_domain = (Get-WmiObject Win32_ComputerSystem).Domain
Write-Verbose "Collection machine: $_collection_machine_name (domain: $_collection_machine_domain))"
$_collection_user_name   = $([Environment]::UserName)
$_collection_user_domain = $([Environment]::UserDomainName)
Write-Verbose "Operator account: $([Environment]::UserDomainName)\$_collection_user_name"
 
# Get the forest context and store them in variables
# Translation table for msDS-Behavior-Version
$_FLPattern = @{
    0 = "DS_BEHAVIOR_WIN2000"
    1 = "DS_BEHAVIOR_WIN2003_WITH_MIXED_DOMAINS"
    2 = "DS_BEHAVIOR_WIN2003"
    3 = "DS_BEHAVIOR_WIN2008"
    4 = "DS_BEHAVIOR_WIN2008R2"
    5 = "DS_BEHAVIOR_WIN2012"
    6 = "DS_BEHAVIOR_WIN2012R2"
}
# Translation table for objectVersion for the schema
$_SchemaPattern = @{
    13 = "Windows 2000 Server"
    30 = "Windows Server 2003"
    31 = "Windows Server 2003 R2"
    44 = "Windows Server 2008"
    47 = "Windows Server 2008 R2"
    56 = "Windows Server 2012"
    69 = "Windows Server 2012 R2"
}
# Searchflags values according to http://msdn.microsoft.com/en-us/library/cc223153.aspx
$_searchFlagsPattern = @{
    0x1 = "fATTINDEX" , "Specifies a hint to the DC to create an index for the attribute."
    0x2 = "fPDNTATTINDEX" , "Specifies a hint to the DC to create an index for the container and the attribute."
    0x4 = "fANR" , "Specifies that the attribute is a member of the ambiguous name resolution (ANR) set."
    0x8 = "fPRESERVEONDELETE" , "Specifies that the attribute MUST be preserved on objects after deletion of the object (that is, when the object is transformed to a tombstone, deleted-object, or recycled-object). This flag is ignored on link attributes, objectCategory, and sAMAccountType."
    0x10 = "fCOPY" , "Specifies a hint to LDAP clients that the attribute is intended to be copied when copying the object. This flag is not interpreted by the server."
    0x20 = "fTUPLEINDEX" , "Specifies a hint for the DC to create a tuple index for the attribute. This will affect the performance of searches where the wildcard appears at the front of the search string."
    0x40 = "fSUBTREEATTINDEX" , "Specifies a hint for the DC to create subtree index for a Virtual List View (VLV) search."
    0x80 = "fCONFIDENTIAL" , "Specifies that the attribute is confidential. An extended access check is required."
    0x100 = "fNEVERVALUEAUDIT" , "Specifies that auditing of changes to individual values contained in this attribute MUST NOT be performed. Auditing is outside of the state model."
    0x200 = "fRODCFilteredAttribute" , "Specifies that the attribute is a member of the filtered attribute set."
    0x400 = "fEXTENDEDLINKTRACKING" , "Specifies a hint to the DC to perform additional implementation-specific, nonvisible tracking of link values. The behavior of this hint is outside the state model."
    0x800 = "fBASEONLY" , "Specifies that the attribute is not to be returned by search operations that are not scoped to a single object. Read operations that would otherwise return an attribute that has this search flag set instead fail with operationsError / ERROR_DS_NON_BASE_SEARCH."
    0x1000 = "fPARTITIONSECRET" , "Specifies that the attribute is a partition secret. An extended access check is required."
}
# Default attributes depending of Schema version
$_default_attributes = @{
    #Schema 2008r2
    47 = @{
        "fPARTITIONSECRET"       = @()
        "fNEVERVALUEAUDIT"       = @("accountExpires","assistant","codePage","company","countryCode","c","department","division","employeeType","homeDirectory","homeDrive","memberOf","localeID","l","logonHours","logonWorkstation","manager","maxStorage","msFVE-RecoveryPassword","msFVE-VolumeGuid","msFVE-KeyPackage","msFVE-RecoveryGuid","msTPM-OwnerInformation","msRADIUS-FramedInterfaceId","msRADIUS-SavedFramedInterfaceId","msRADIUS-FramedIpv6Prefix","msRADIUS-SavedFramedIpv6Prefix","msRADIUS-FramedIpv6Route","msRADIUS-SavedFramedIpv6Route","msNPAllowDialin","msNPCallingStationID","msNPSavedCallingStationID","msRADIUSCallbackNumber","msRADIUSFramedIPAddress","msRADIUSFramedRoute","msRADIUSServiceType","msRASSavedCallbackNumber","msRASSavedFramedIPAddress","msRASSavedFramedRoute","otherLoginWorkstations","postOfficeBox","postalAddress","postalCode","preferredOU","primaryGroupID","profilePath")
        "fCONFIDENTIAL"          = @("msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKI-CredentialRoamingTokens","msPKIRoamingTimeStamp","msPKIDPAPIMasterKeys","msPKIAccountCredentials")
        "fRODCFilteredAttribute" = @("fRODCFilteredAttribute","msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKIRoamingTimeStamp","msPKIDPAPIMasterKeys","msPKIAccountCredentials")
    }
    #Schema 2012
    56 = @{
        "fCONFIDENTIAL"          = @("msKds-Version","msKds-DomainID","msKds-KDFParam","msKds-CreateTime","msKds-RootKeyData","msKds-UseStartTime","msKds-KDFAlgorithmID","msKds-PublicKeyLength","msKds-PrivateKeyLength","msKds-SecretAgreementParam","msKds-SecretAgreementAlgorithmID","msDS-TransformationRulesCompiled","msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKI-CredentialRoamingTokens","msPKI-RoamingTimeStamp","msPKI-DPAPIMasterKeys","msPKI-AccountCredentials","UnixUserPassword","msTPM-OwnerInformationTemp")
        "fRODCFilteredAttribute" = @("msKds-Version","msKds-DomainID","msKds-KDFParam","msKds-CreateTime","msKds-RootKeyData","msKds-UseStartTime","msKds-KDFAlgorithmID","msKds-PublicKeyLength","msKds-PrivateKeyLength","msKds-SecretAgreementParam","msKds-SecretAgreementAlgorithmID","msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKI-RoamingTimeStamp","msPKI-DPAPIMasterKeys","msPKI-AccountCredentials","msTPM-OwnerInformationTemp")
    
    }
    #Schema 2012r2
    69 = @{
        "fCONFIDENTIAL"          = @("msKds-Version","msKds-DomainID","msKds-KDFParam","msKds-CreateTime","msKds-RootKeyData","msKds-UseStartTime","msKds-KDFAlgorithmID","msKds-PublicKeyLength","msKds-PrivateKeyLength","msKds-SecretAgreementParam","msKds-SecretAgreementAlgorithmID","msDS-TransformationRulesCompiled","msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKI-CredentialRoamingTokens","msPKI-RoamingTimeStamp","msDS-IssuerCertificates","msPKI-DPAPIMasterKeys","msPKI-AccountCredentials","UnixUserPassword","msTPM-OwnerInformationTemp")
        "fRODCFilteredAttribute" = @("msKds-Version","msKds-DomainID","msKds-KDFParam","msKds-CreateTime","msKds-RootKeyData","msKds-UseStartTime","msKds-KDFAlgorithmID","msKds-PublicKeyLength","msKds-PrivateKeyLength","msKds-SecretAgreementParam","msKds-SecretAgreementAlgorithmID","msFVE-RecoveryPassword","msFVE-KeyPackage","msTPM-OwnerInformation","msPKI-RoamingTimeStamp","msPKI-DPAPIMasterKeys","msPKI-AccountCredentials","msTPM-OwnerInformationTemp")
    }
}
# Checking if we are going to work on the current forest or if we target another one
If ( $Forest -eq $null -or $Forest -eq "" )
{
    Write-Debug   "Entering [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()"
    $_forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
} Else {
    Write-Verbose "A specific forest has been specifed in the input paramters"
    Try
    {
        Write-Debug "Invoquing an object System.DirectoryServices.ActiveDirectory.DirectoryContext"
        $_forest_context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Forest",$Forest)
        Write-Debug "Builing a System.DirectoryServices.ActiveDirectory.Forest object for $Forest"
        $_forest         = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest( $_forest_context )
    }
    Catch
    {
        Write-Error "Cannot access the forest $Forest"
        Throw "Cannot access the forest $Forest"
    }
}
Write-Debug "Gathering forest information"
$_forest_name         = $($_forest.RootDomain.Name)
# We do not store $_forest.Domains into a variable because it is time consuming
#$_forest_domains      = $_forest.Domains
$_forest_config       = $(([ADSI]"LDAP://RootDSE").configurationNamingContext)
$_forest_schema       = $(([ADSI]"LDAP://RootDSE").schemaNamingContext)
$_forest_mode         = $($_forest.ForestMode)
$_forest_schema_level = $(([ADSI]"LDAP://$_forest_schema").objectVersion)
$_forest_root         = $($_forest.RootDomain.Name)
$_forest_creation_time= $(([ADSI]"LDAP://$_forest_schema").whenCreated)
# The dSHeuristics can influence the adminSDHolder, so we look at it
$_forest_dSHeuristics = $(([ADSI]"LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($_forest_config)").dSHeuristics)
$_admin_protected_ao = $true
$_admin_protected_bo = $true
$_admin_protected_so = $true
$_admin_protected_po = $true
If ( $_forest_dSHeuristics -ne $null )
{
    $script:_issues += @{
        "Title"    = "The dsHeuristic has been changed (value=$_forest_dSHeuristics)"
        "Severity" = 1
    }
    If ( $_forest_dSHeuristics[15] -ne $null) {
        #Parse ds
        $_dsHeuristicPattern = @{
            0x1 = "Account Operators" 
            0x2 = "Server Operators"
            0x4 = "Print Operators"
            0x8 = "Backup Operators"
        }
        $_dsHeuristicPattern.Keys | Where-Object { $_ -band $_forest_dSHeuristics[15] } | ForEach-Object {
            $_excluded_group_name  = $_dsHeuristicPattern.Item($_)
            $_excluded_group_index = $_
            Write-Verbose "$_excluded_group_name is excluded from the adminSDHolder protection"
            Switch ( $_excluded_group_index )
            {
                0x1 { $_admin_protected_ao = $false ; $_issues += @{ "Title" = "The default adminSDHolder is modified through the dSHeuristics" ; "Objects" = "Account Operators group is protected by the adminSDHolder" ; "Severity" = 2} }
                0x2 { $_admin_protected_so = $false ; $_issues += @{ "Title" = "The default adminSDHolder is modified through the dSHeuristics" ; "Objects" = "Server Operators group is protected by the adminSDHolder" ; "Severity" = 2} }
                0x4 { $_admin_protected_po = $false ; $_issues += @{ "Title" = "The default adminSDHolder is modified through the dSHeuristics" ; "Objects" = "Print Operators group is protected by the adminSDHolder" ; "Severity" = 2} }
                0x8 { $_admin_protected_bo = $false ; $_issues += @{ "Title" = "The default adminSDHolder is modified through the dSHeuristics" ; "Objects" = "Backup Operators group is protected by the adminSDHolder" ; "Severity" = 2} }
            }
        }
    }
}
Write-Debug "Account Operators protected: $_admin_protected_ao"
Write-Debug "Server Operators protected: $_admin_protected_so"
Write-Debug "Print Operators protected: $_admin_protected_po"
Write-Debug "Backup Operators protected: $_admin_protected_bo"
# Get the tombstoneLifetime for further use
$_forest_TSL = $(([ADSI]"LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($_forest_config)").tombstoneLifetime)
# If the tombstoneLifetime isnt defined, we set it to its default value, 60
#How many time and when was the last time the dsheuristic was changed as well and the TSL
Write-Debug "Invoquing an object System.DirectoryServices.ActiveDirectory.DirectoryContext"
$_forest_meta_context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$_forest_name)
# .NET Class that returns a Domain Controller for Specified Context
Write-Debug   "Calling FindOne()"
# This object will be stored in the ht_table to be able to ask for replication metadata later on the script
$_forest_meta_loaded = [System.DirectoryServices.ActiveDirectory.DomainController]::findOne($_forest_meta_context)
Write-Debug "Calling GetReplicationMetadata(...) for the schema information"
$_object_meta = $_forest_meta_loaded.GetReplicationMetadata($_forest_schema)
Try   { $_forest_schema_version = $($_object_meta.objectversion.Version) }
Catch { $_forest_schema_version = "N/A" }
Try   { $_forest_schema_whenModified = $( ($_object_meta.objectversion.LastOriginatingChangeTime).ToUniversalTime() ) }
Catch { $_forest_schema_whenModified = "N/A" }
Write-Debug "Calling GetReplicationMetadata(...) for the configuration information"
$_object_meta = $_forest_meta_loaded.GetReplicationMetadata($("CN=Directory Service,CN=Windows NT,CN=Services,$($_forest_config)"))
Try   { $_forest_TSL_version = $($_object_meta.tombstonelifetime.Version) }
Catch { $_forest_TSL_version = "N/A" }
Try   { $_forest_TSL_whenModified = $( ($_object_meta.tombstonelifetime.LastOriginatingChangeTime).ToUniversalTime() ) }
Catch { $_forest_TSL_whenModified = "N/A" }
Try   { $_forest_dSHeuristics_version = $($_object_meta.dsheuristics.Version) } 
Catch { $_forest_dSHeuristics_version = "N/A" }
Try   { $_forest_dSHeuristics_whenModified = $( ($_object_meta.dsheuristics.LastOriginatingChangeTime).ToUniversalTime() ) }
Catch { $_forest_dSHeuristics_whenModified = "N/A" }
Write-Verbose "Forest name: $_forest_name"
Write-Debug   "Forest root: $_forest_root"
Write-Debug   "Forest mode: $_forest_mode ($($_FLPattern[ $_forest_mode ]))"
Write-Debug   "Forest creation: $_forest_creation_time"
Write-Debug   "Configuration DN: $_forest_config"
Write-Debug   "Schema DN: $_forest_schema"
Write-Debug   "Schema level: $_forest_schema_level ($($_SchemaPattern[$_forest_schema_level]))"
Write-Debug   "Schema version: $_forest_schema_version"
Write-Debug   "Schema last modified: $_forest_schema_whenModified"
Write-Verbose "dSHeuristics: $_forest_dSHeuristics"
Write-Debug   "dSHeuristics version: $_forest_dSHeuristics_version"
Write-Debug   "dSHeuristics last modified: $_forest_dSHeuristics_whenModified"
If ( $_forest_TSL -eq $null -or $_forest_TSL -eq "")
{
    Write-Debug "tombstoneLifetime is not set, default is 60"
    $_forest_TSL = 60
}
Write-Debug   "tombstoneLifetime: $_forest_TSL"
Write-Debug   "tombstoneLifetime version: $_forest_TSL_version"
Write-Debug   "tombstoneLifetime last modified: $_forest_TSL_whenModified"
Write-Verbose "--"
# Initialize the variable in case of a SkipSchema option enable 
$_ht_schema_classes = "N/A"
$_ht_schema_attributes = "N/A"
# Check if the operator wants to skip the schema export, that's legit, there is no analysis of it so far
If ( $SkipSchema -eq $false )
{
    # Collect all classes default security descriptors
    Write-Verbose "List all classes and their associated security descriptors"
    $_ht_schema_classes = @{}
    Write-Debug   "Creating an object System.DirectoryServices.DirectorySearcher"
    $_all_classes = New-Object System.DirectoryServices.DirectorySearcher
    $_all_classes.SearchRoot = "LDAP://$_forest_schema"
    $_all_classes.Filter = "(objectCategory=CN=Class-Schema,$_forest_schema)"
    $_all_classes.PageSize = 1000
    $_all_classes.SearchScope = "Subtree"
    Write-Debug   "Call FindAll()"
    $_i_classes = 0
    $_all_classes_res = $_all_classes.FindAll()
    $_i_classes_total = $_all_classes_res.Count
    $_all_classes_res | ForEach-Object `
    {
        Write-Progress -Activity "Collecting classes' defaultSecurityDescriptor .." -Status "$($_i_classes+1) / $($_i_classes_total)" -CurrentOperation "Collecting class $($_.GetDirectoryEntry().lDAPDisplayName).." -Id 0 -PercentComplete ( $_i_classes++ / $_i_classes_total * 100 ) 
        Write-Debug "Calling GetReplicationMetadata($($_.GetDirectoryEntry().lDAPDisplayName)) for the defaultSecurityDescriptor information"
        $_object_meta = $_forest_meta_loaded.GetReplicationMetadata($($_.GetDirectoryEntry().distinguishedName))
        Try   { $_schema_defaultsd_version = $($_object_meta.defaultsecuritydescriptor.Version) }
        Catch { $_schema_defaultsd_version = $_forest_schema_version = "N/A" }
        Try   { $_schema_defaultsd_last_modified = $( ($_object_meta.defaultsecuritydescriptor.LastOriginatingChangeTime).ToUniversalTime() ) }
        Catch { $_schema_defaultsd_last_modified = "N/A" }
 
        $_ht_schema_classes += @{
            $($_.GetDirectoryEntry().lDAPDisplayName) = @{
                "defaultSecurityDescriptor" = $($_.GetDirectoryEntry().defaultSecurityDescriptor)
                "defaultSecurityDescriptorVersion" = $_schema_defaultsd_version
                "defaultSecurityDescriptorWhenModified" = $_schema_defaultsd_last_modified
            }
        }
    }
    # Should also check for schemaFlagsEx Flags
    # http://msdn.microsoft.com/en-us/library/cc223872.aspx CR (FLAG_ATTR_IS_CRITICAL, 0x00000001)
    # Specifies that the attribute is not a member of the filtered attribute set even if the fRODCFilteredAttribute flag is set. For more information, see sections 3.1.1.2.3 and 3.1.1.2.3.5.
    # Used to keep track of security related searchFlags settings
    $_ht_check_attributes = @{ "fNEVERVALUEAUDIT" = @() ; "fRODCFilteredAttribute" = @() ; "fPARTITIONSECRET" = @() ; "fCONFIDENTIAL" = @() ; "FLAG_ATTR_IS_CRITICAL" = @() }
    # Collect all attributes searchflags
    # I know, we could have done that in one DirectorySearcher, it needs to be recoded
    #! Recode that part to do everything (classes and attributes) in the same DirectorySearcher
    Write-Verbose "List all attributes and their associated searchflags"
    $_ht_schema_attributes = @{}
    Write-Debug   "Creating an object System.DirectoryServices.DirectorySearcher"
    $_all_attributes = New-Object System.DirectoryServices.DirectorySearcher
    $_all_attributes.SearchRoot = "LDAP://$_forest_schema"
    $_all_attributes.Filter = "(objectCategory=CN=Attribute-Schema,$_forest_schema)"
    $_all_attributes.PageSize = 1000
    $_all_attributes.SearchScope = "Subtree"
    Write-Debug   "Call FindAll()"
    $_i_attributes = 0
    $_all_attributes_res = $_all_attributes.FindAll()
    $_i_attributes_total = $_all_attributes_res.Count
    $_all_attributes_res | ForEach-Object `
    {
        $_current_attribute = $($_.GetDirectoryEntry().lDAPDisplayName)
        Write-Progress -Activity "Collecting attributes' searchFlags.." -Status "$($_i_attributes+1) / $($_i_attributes_total)" -CurrentOperation "Collecting class $_current_attribute.." -Id 0 -PercentComplete ( $_i_attributes++ / $_i_attributes_total * 100 ) 
        Write-Debug "Calling GetReplicationMetadata($_current_attribute) for the searchFlags information"
        $_object_meta = $_forest_meta_loaded.GetReplicationMetadata($($_.GetDirectoryEntry().distinguishedName))
        $_current_searchFlags = $($_.GetDirectoryEntry().searchFlags)
        # Do the same thing for the attribute searchFlags
        Try   { $_schema_searchflags_version = $($_object_meta.searchflags.Version) }
        Catch { $_schema_searchflags_version = "N/A" }
        Try   { $_schema_searchflags_last_modified = $( ($_object_meta.searchflags.LastOriginatingChangeTime).ToUniversalTime() ) }
        Catch { $_schema_searchflags_last_modified = "N/A" }
 
        $_ht_schema_attributes += @{
            $($_.GetDirectoryEntry().lDAPDisplayName) = @{
                "searchFlags" = $_current_searchFlags
                "searchFlagsVersion" = $_schema_searchflags_version
                "searchFlagsWhenModified" = $_schema_searchflags_last_modified
                "schemaFlagsEx" = $($_.GetDirectoryEntry().schemaFlagsEx)
            }
        }
        # Store attributes with a security releated searchFlags enabled
        If ( $_current_searchFlags -band 0x80 )   { $_ht_check_attributes["fCONFIDENTIAL"] += $_current_attribute }
        If ( $_current_searchFlags -band 0x200 )  { $_ht_check_attributes["fRODCFilteredAttribute"] += $_current_attribute }
        If ( $_current_searchFlags -band 0x10 )   { $_ht_check_attributes["fNEVERVALUEAUDIT"] += $_current_attribute }
        If ( $_current_searchFlags -band 0x1000 ) { $_ht_check_attributes["fPARTITIONSECRET"] += $_current_attribute }
        If ( $($_.GetDirectoryEntry().schemaFlagsEx) -band 0x1 ) { $_ht_check_attributes["FLAG_ATTR_IS_CRITICAL"] += $_current_attribute }
    }
} # End of the SkipSchema check
$script:_ht_forest    = @{
    "Name" = $_forest_name
    "Mode" = $($_forest.ForestMode)
    "SchemaLevel" = $_forest_schema_level
    "SchemaLevelVersion" = $_forest_schema_version
    "SchemaLevelWhenModified" = $_forest_schema_whenModified
    "ForestdSHeuristics" = $_forest_dSHeuristics
    "ForestdSHeuristicsVersion" = $_forest_dSHeuristics_version
    "ForestdSHeuristicsWhenModified" = $_forest_dSHeuristics_whenModified
    "TSL" = $_forest_TSL
    "TSLVersion" = $_forest_TSL_version
    "TSLWhenModified" = $_forest_TSL_whenModified
    "SchemaSecurity" = $_ht_schema_classes
    "SchemaAttributes" = $_ht_schema_attributes
    "SensitiveAttributes" = $_ht_check_attributes
    "CreationTime" = $_forest_creation_time
    "CollectionMachineName" = $_collection_machine_name
    "CollectionMachineDomain" = $_collection_machine_domain
    "CollectionUserName" = $_collection_user_name
    "CollectionUserDomain" = $_collection_user_domain
    "CollectionTimestamp" = $_timestamp
}
$_i_domains = 0 # This is used to show a progress bar
$_forest.Domains | ForEach-Object `
{
    Write-Progress -Activity "Gathering domain information.." -Status "$($_i_domains+1) / $($_forest.Domains.Count)" -CurrentOperation "Collecting information for the domain $($_.Name).." -Id 0 -PercentComplete ( $_i_domains++ / $_forest.Domains.Count * 100 )
    Write-Debug "Gathering information for the domain: $($_.name)"
    # Get some info about the current domain we parse
    $_domain_dns = $_.Name
    $_domain_dns_parent = $_.Parent.Name
    $_domain = [ADSI]"LDAP://$_domain_dns"
    $_domain_dns_dc = $_.Name  # For now... It is better to find an actual DC
    $_domain_sid = (New-Object Security.Principal.Securityidentifier($_domain.objectSid[0],0)).Value # Get the SID to a S-x-Y... format
    $_domain_netbios = $($_domain.name) # Get the NetBIOS name
    $_domain_dn = $($_domain.distinguishedName) # Get the DN of the domain
    $_domain_fl = $($_domain."msDS-Behavior-Version")
    $_domain_creation_time = $($_domain.whenCreated)
    $_domain_Admin500 = $(([ADSI]"LDAP://$_domain_dns_dc/<SID=$_domain_sid-500>").sAMAccountName)
    $_domain_Admin500_UAC = $(([ADSI]"LDAP://$_domain_dns_dc/<SID=$_domain_sid-500>").userAccountControl)
    $_domain_Guest501 = $(([ADSI]"LDAP://$_domain_dns_dc/<SID=$_domain_sid-501>").sAMAccountName)
    $_domain_Guest501_UAC = $(([ADSI]"LDAP://$_domain_dns_dc/<SID=$_domain_sid-501>").userAccountControl)
    Write-Debug   "DNS: $_domain_dns"
    Write-Debug   "Parent domain: $_domain_dns_parent"
    Write-Debug   "Creation Time: $_domain_creation_time"
    Write-Verbose "DC: $_domain_dns_dc"
    Write-Verbose "SID: $_domain_sid"
    Write-Debug   "NetBIOS: $_domain_netbios"
    Write-Debug   "DN: $_domain_dn"
    Write-Debug   "FL: $_domain_fl  ($($_FLPattern[$_domain_fl]))"
    Write-Verbose "Admin (SID-500): $_domain_Admin500"
    Write-Debug   "Admin UAC (SID-500): $_domain_Admin500_UAC"
    Write-Verbose "Guest (SID-501): $_domain_Guest501"
    Write-Debug   "Guest UAC (SID-501): $_domain_Guest501_UAC"
    $_domain_quota = @{
        "ms-DS-MachineAccountQuota"        = $($_domain."ms-DS-MachineAccountQuota") 
        "msDS-AllUsersTrustQuota"          = $($_domain."msDS-AllUsersTrustQuota")
        "msDS-PerUserTrustQuota"           = $($_domain."msDS-PerUserTrustQuota")
        "msDS-PerUserTrustTombstonesQuota" = $($_domain."msDS-PerUserTrustTombstonesQuota")
    }
    Write-Debug  "Quota ms-DS-MachineAccountQuota: $($_domain_quota["ms-DS-MachineAccountQuota"])"
    Write-Debug  "Quota msDS-AllUsersTrustQuota: $($_domain_quota["msDS-AllUsersTrustQuota"])"
    Write-Debug  "Quota msDS-PerUserTrustQuota: $($_domain_quota["msDS-PerUserTrustQuota"])"
    Write-Debug  "Quota msDS-PerUserTrustTombstonesQuota: $($_domain_quota["msDS-PerUserTrustTombstonesQuota"])"
 
    # Check if the operator wants to skip the trust listing, I would not advice that, but it fasten the process on corp ;)
    If ( $SkipTrust -eq $false )
    {
        # List all external trusts
        $_ht_sid_trust = @{}
        Write-Verbose "List all external trusts"
        Write-Debug   "Creating an object System.DirectoryServices.DirectorySearcher"
        $_domain_trust = New-Object System.DirectoryServices.DirectorySearcher
        $_domain_trust.SearchRoot = "LDAP://$_domain_dns_dc/CN=System,$_domain_dn"
        $_domain_trust.Filter = "(&(objectCategory=trustedDomain)(!trustAttributes:1.2.840.113556.1.4.803:=32))" # Get the external trusts
        $_domain_trust.SearchScope = "Subtree"
        Write-Debug   "Call FindAll()"
        $_domain_trust = $_domain_trust.FindAll()
        $_i_trust_total = $_domain_trust.Count # use for the progress bar
        $_i_trust_inc   = 0                    # use for the progress bar
        # Parse each of the external trust we found
        $_domain_trust | ForEach-Object `
        {
            # Get the SID of the trusted domain
            Try
            {
                $_domain_trust_sid = (New-Object Security.Principal.Securityidentifier(($_.GetDirectoryEntry()).securityIdentifier[0],0)).Value # Get the SID
            }
            Catch
            {
                $_domain_trust_sid = "N/A"
                Write-Verbose "ERROR: Cannot get the SID of the trsuted domain"
            }
            Write-Verbose "Trust SID: $_domain_trust_sid"
            Try
            {
                $_domain_trust_flatName        = $(($_.GetDirectoryEntry()).flatName)
                Write-Progress -Activity "Gathering trust information for the domain $_domain_dns.." -Status "$($_i_trust_inc+1) / $_i_trust_total" -CurrentOperation "Collecting information for the trust with $_domain_trust_flatName.." -Id 1 -ParentId 0 -PercentComplete ( $_i_trust_inc++ / $_i_trust_total * 100 )
                $_domain_trust_trustPartner    = $(($_.GetDirectoryEntry()).trustPartner)
                $_domain_trust_trustAttributes = $(($_.GetDirectoryEntry()).trustAttributes)
                $_domain_trust_trustDirection  = $(($_.GetDirectoryEntry()).trustDirection)
                $_domain_trust_whenCreated     = $(($_.GetDirectoryEntry()).whenCreated)
            }
            Catch
            {
                $_domain_trust_flatName        = "N/A"
                $_domain_trust_trustPartner    = "N/A"
                $_domain_trust_trustAttributes = "N/A"
                $_domain_trust_trustDirection  = "N/A"
                $_domain_trust_whenCreated     = "N/A"
                Write-Verbose "ERROR: Cannot get info from the trusted domain with GetDirectoryEntry()"
            }
            Write-Verbose "Trust flat name: $_domain_trust_flatName"
            Write-Debug   "trustPartner: $_domain_trust_trustPartner"
            Write-Debug   "trustAttributes: $_domain_trust_trustAttributes"
            Write-Debug   "trustDirection: $_domain_trust_trustDirection"
            Write-Debug   "whenCreated: $_domain_trust_whenCreated"
         
            # Get the last password change.
            #! NOTE THIS NEEDS TO BE REVIEWED, flatName$ accounts are not always there
            $_domain_trust_pwdLastSet = ([ADSI]"LDAP://$_domain_dns_dc/CN=$_domain_trust_flatName$,CN=Users,$_domain_dn")
            # Convert the datetime NTTE thingy into something we can understand
            Try {
                Write-Debug "Try to convert the pwdLastSet to a readable format"
                $_domain_trust_pwdLastSet = [datetime]::fromfiletime( $_domain_trust_pwdLastSet.ConvertLargeIntegerToInt64( $_domain_trust_pwdLastSet.pwdLastSet.value ))
                $_domain_trust_pwdLastSet = $_domain_trust_pwdLastSet.ToUniversalTime()
            }
            Catch {
                $_domain_trust_pwdLastSet = "N/A"
                Write-Debug "Failed to convert the pwdLastSet into a readable format. Set N/A instead"
            }
            Write-Debug   "pwdLastSet: $_domain_trust_pwdLastSet"
            # Prepare an object containing the trust info
            $_trust_item = @{
                flatName        = $_domain_trust_flatName
                trustAttributes = $_domain_trust_trustAttributes
                trustDirection  = $_domain_trust_trustDirection
                trustPartner    = $_domain_trust_trustPartner
                whenCreated     = $_domain_trust_whenCreated
                pwdLastSet      = $_domain_trust_pwdLastSet
            }
            # Add the trust to the hashtable
            Try   { $_ht_sid_trust += @{ $_domain_trust_sid = $_trust_item } }
            # We avoid error message because of duplicate key for the _ht_sid_trust
            Catch { Write-Debug "The trust $_domain_trust_sid is aleady in this hashtable" }
            Write-Verbose "--"
        }
    } # End of the SkipTrust check
 
    # Sets Context to Domain for System.DirectoryServices.ActiveDirectory.DomainController
    Write-Verbose "Creating a DirectoryContext object to ba able to query replication's metadata on that domain"
    Write-Debug   "Create an object System.DirectoryServices.ActiveDirectory.DirectoryContext for the iterated domain"
    $_domain_meta_context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("Domain",$_domain_dns_dc)
    # .NET Class that returns a Domain Controller for Specified Context
    Write-Debug   "Calling FindOne()"
    # This object will be stored in the ht_table to be able to ask for replication metadata later on the script
    $_domain_meta = [System.DirectoryServices.ActiveDirectory.DomainController]::findOne($_domain_meta_context)
 
    # Build the list of group to query
    # Well-known security identifiers in Windows operating systems https://support.microsoft.com/kb/243330
    Write-Debug  "Building the group list"
    $_ht_groups = @{}
    # We add a special scnenario if the domain is the root domain
    If ( $_domain_dns -eq $_forest_root )
    {
        # Schema Admins
        $_group_item = @{
            SID = "$_domain_sid-518"
            AdminSDHolderProtected = 1
            ActualName = ""
            ActualDN = ""
            ActualMembers =  @()
            Admins = $true
            LimitWarning = 0
            LimitError   = 0
        }
        $_ht_groups += @{ "Schema Admins" = $_group_item }
        # Enterprise Admins
        $_group_item = @{
            SID = "$_domain_sid-519"
            AdminSDHolderProtected = 1
            ActualName = ""
            ActualDN = ""
            ActualMembers =  @()
            Admins = $true
            LimitWarning = 2
            LimitError   = 5
        }
        $_ht_groups += @{ "Enterprise Admins" = $_group_item }
    }
    # Administrators
    $_group_item = @{
        SID = "S-1-5-32-544"
        AdminSDHolderProtected = 1
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 5
        LimitError   = 5
    }
    $_ht_groups += @{"Administrators" = $_group_item }
    # Print Operators
    $_group_item = @{
        SID = "S-1-5-32-550"
        AdminSDHolderProtected = $_admin_protected_po
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Print Operators" = $_group_item }
    # Remote Desktop Users
    $_group_item = @{
        SID = "S-1-5-32-555"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Remote Desktop Users" = $_group_item }
    # Account Operators
    $_group_item = @{
        SID = "S-1-5-32-548"
        AdminSDHolderProtected = $_admin_protected_ao
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Account Operators" = $_group_item }
    # Backup Operators
    $_group_item = @{
        SID = "S-1-5-32-551"
        AdminSDHolderProtected = $_admin_protected_po
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Backup Operators" = $_group_item }
    # Cert Publishers
    $_group_item = @{
        SID = "$_domain_sid-517"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = -1
        LimitError   = -1
    }
    $_ht_groups += @{"Cert Publishers" = $_group_item }
    # Domain Admins
    $_group_item = @{
        SID = "$_domain_sid-512"
        AdminSDHolderProtected = 1
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 5
        LimitError   = 10
    }
    $_ht_groups += @{"Domain Admins" = $_group_item }
    # Domain Controllers
    $_group_item = @{
        SID = "$_domain_sid-516"
        AdminSDHolderProtected = 1
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = -1
        LimitError   = -1
    }
    $_ht_groups += @{"Domain Controllers" = $_group_item }
    # Read-only Domain Controllers
    $_group_item = @{
        SID = "$_domain_sid-521"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = -1
        LimitError   = -1
    }
    $_ht_groups += @{"Read-only Domain Controllers" = $_group_item }
    # Server Operators
    $_group_item = @{
        SID = "S-1-5-32-549"
        AdminSDHolderProtected = $_admin_protected_so
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Server Operators" = $_group_item }
    # Domain Guests
    $_group_item = @{
        SID = "$_domain_sid-514"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = 1
        LimitError   = 2
    }
    $_ht_groups += @{"Domain Guests" = $_group_item }
    # Group Policy Creators Owners
    $_group_item = @{
        SID = "$_domain_sid-520"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $true
        LimitWarning = 1
        LimitError   = 2
    }
    $_ht_groups += @{"Group Policy Creator Owners" = $_group_item }
    # Guests
    $_group_item = @{
        SID = "S-1-5-32-546"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Guests" = $_group_item }
    # Windows Authorization Access Group
    <#$_group_item = New-Object PSObject -Property @{
        SID = "S-1-5-32-560"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
    }
    $_ht_groups += @{"Windows Authorization Access Group" = $_group_item }
    #>
    # RAS and IAS Servers
    $_group_item = @{
        SID = "$_domain_sid-553"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"RAS and IAS Servers" = $_group_item }
    # Pre-Windows 2000 Compatible Access
    $_group_item = @{
        SID = "S-1-5-32-554"
        AdminSDHolderProtected = 0
        ActualName = ""
        ActualDN = ""
        ActualMembers =  @()
        Admins = $false
        LimitWarning = 0
        LimitError   = 0
    }
    $_ht_groups += @{"Pre-Windows 2000 Compatible Access" = $_group_item }
    # Add the check of custom groups (e.g. DNSAdmins but could also be DHCP...)
    $_current_domain_custom_group = New-Object System.DirectoryServices.DirectorySearcher
    $_current_domain_custom_group_filter = "(&(objectCategory=Group)(name=DNSAdmins))"
    $_current_domain_custom_group.SearchRoot = "LDAP://$_domain_dns_dc"
    $_current_domain_custom_group.Filter = $_current_domain_custom_group_filter
    $_current_domain_custom_group.SearchScope = "Subtree"
    $_current_domain_custom_group_result = $_current_domain_custom_group.FindOne()
    # It it exist, we add it to $_current_domain_groups
    If ( $_current_domain_custom_group_result ) {
        # DNSAdmins
        Write-Debug "Adding DNSAdmins group to the collection for that domain."
        $_group_item = @{
            SID = $((New-Object Security.Principal.Securityidentifier(($_current_domain_custom_group_result.GetDirectoryEntry()).objectSid[0],0)).Value )
            AdminSDHolderProtected = 0
            ActualName = ""
            ActualDN = ""
            ActualMembers =  @()
            Admins = $true
            LimitWarning = 2
            LimitError   = 2
        }
        $_ht_groups += @{"DNSAdmins" = $_group_item }
    }
    # Manage custom groups
    If ( $AddGroups -ne $null)
    {
        Write-Verbose "Custom groups collection is set"
        $AddGroups | ForEach-Object `
        {
            $_current_custom_group = $_
            Write-Verbose "Checking if $_current_custom_group is on $_domain_netbios"
            $_current_custom_group_domain = $_current_custom_group.Split("\")[0]
            $_current_custom_group_name   = $_current_custom_group.Split("\")[1]
            If ( $_current_custom_group_domain.ToUpper() -eq $_domain_netbios.ToUpper() )
            {
                Write-Verbose "$_current_custom_group is on $_domain_netbios, adding it to the collection"
                $_current_domain_custom_group = New-Object System.DirectoryServices.DirectorySearcher
                $_current_domain_custom_group_filter = "(&(objectCategory=Group)(sAMAccountName=$_current_custom_group_name))"
                $_current_domain_custom_group.SearchRoot = "LDAP://$_domain_dns_dc"
                $_current_domain_custom_group.Filter = $_current_domain_custom_group_filter
                $_current_domain_custom_group.SearchScope = "Subtree"
                $_current_domain_custom_group_result = $_current_domain_custom_group.FindOne()
                # It it exist, we add it to $_current_domain_groups
                If ( $_current_domain_custom_group_result ) {                    
                    Write-Debug "Adding $_current_custom_group_name group to the collection for that domain."
                    $_group_item = @{
                        SID = $((New-Object Security.Principal.Securityidentifier(($_current_domain_custom_group_result.GetDirectoryEntry()).objectSid[0],0)).Value )
                        AdminSDHolderProtected = 0
                        ActualName = ""
                        ActualDN = ""
                        ActualMembers =  @()
                        Admins = $true
                        LimitWarning = -1
                        LimitError   = -1
                    }
                    $_ht_groups += @{ $_current_custom_group_name = $_group_item }
                } Else {
                    Write-Verbose "The custom group $_current_custom_group has not been found."
                }
            }
        }
    }
    $_item_domain = @{
        DomainDC          = $_domain_dns_dc
        DomainDN          = $_domain_dn
        DomainName        = $_domain_dns
        DomainParentName  = $_domain_dns_parent
        DomainNetbios     = $_domain_netbios
        DomainAdmin500    = $_domain_Admin500
        DomainAdmin500UAC = $_domain_Admin500_UAC
        DomainGuest501    = $_domain_Guest501
        DomainGuest501UAC = $_domain_Guest501_UAC
        DomainFL          = $_domain_fl
        DomainQuota       = $_domain_quota
        DomainGroups      = $_ht_groups
        DomainTrusts      = $_ht_sid_trust
        DomainMeta        = $_domain_meta
        CreationTime      = $_domain_creation_time
    }
    $script:_ht_sid_domain += @{ $_domain_sid = $_item_domain }
    Write-Verbose "--"
}
 
# hastable with the well-known SID that we are looking for
$_wellKnownSecurity = @{
    "S-1-1-0" = "Everyone" ; 
    "S-1-3-0" = "Creator Owner" ;
    "S-1-3-1" = "Creator Group" ;
    "S-1-5-10" = "Self" ;
    "S-1-5-11" = "Authenticated Users" ;
    "S-1-5-2" = "Network" ;
    "S-1-5-3" = "Batch" ;
    "S-1-5-4" = "Interactive" ;
    "S-1-5-6" = "Service" ;
    "S-1-5-7" = "Anonymous Logon" ;
    "S-1-5-1" = "Dialup" ;
    "S-1-5-8" = "Proxy" ;
    "S-1-5-9" = "Enterprise Domain Controllers" ;
    "S-1-5-12" = "Restricted" ;
    "S-1-5-18" = "Well-Known-Security-Id-System" ;
    "S-1-5-13" = "Terminal Server User" ;
    "S-1-5-19" = "Local Service" ;
    "S-1-5-20" = "Network Service" ;
    "S-1-5-15" = "This Organization" ;
    "S-1-5-1000" = "Other Organization" ;
    "S-1-5-14" = "Remote Interactive Logon" ;
    "S-1-5-64-10" = "NTLM Authentication" ;
    "S-1-5-64-21" = "Digest Authentication" ;
    "S-1-5-64-14" = "SChannel Authentication"
}
 
 
function _IncludeObjects()
{
    param (
        [string] $_SID_to_add,
        [string] $_Type_to_add
    )
    Write-Debug "++ _IncludeObjects"
    Write-Debug "_SID_to_add: $_SID_to_add"
    Write-Debug "_Type_to_add: $_Type_to_add"
    # We first check if the object SID isnt already in the hashtable
    If ( $script:_ht_sid_objects.Keys -notcontains $_SID_to_add )
    {
        Write-Debug "First time we see this object, $_SID_to_add will be added to _ht_sid_objects"
        # It is not so we add it
        # Depending on the type, we perform additional query to get more info about the object
        #! This version of the script doesnt collect object info now.
        
        $script:_ht_sid_objects += @{ $_SID_to_add = @{ "type" = $_Type_to_add ; "AdminSDHolderProtected" = $script:_current_group_AdminSDHolderProtected ; "membership" = @( $script:_current_group_dn_inloop ) ; "Admin" = $script:_current_group_Admins } }
    } Else {
 
        Write-Debug "This object $_SID_to_add is already in _ht_sid_objects"
        # The following lines will help to determine admincount orphans.
        # We check if it is already an object member of an AdminSDHolderProtected group
        If ( $script:_ht_sid_objects[ $_SID_to_add ].AdminSDHolderProtected -eq 0 -and $script:_current_group_AdminSDHolderProtected -eq 1 )
        {
            #Then we set the user as a member of a protected group
            $script:_ht_sid_objects[ $_SID_to_add ].AdminSDHolderProtected = 1
        }
        If ( $script:_ht_sid_objects[ $_SID_to_add ].Admins -eq $false -and $script:_current_group_Admins -eq $true )
        {
            #Then we set the user as a member of a protected group
            $script:_ht_sid_objects[ $_SID_to_add ].Admins = $true
        }
        # In any case, add the membership
        $script:_ht_sid_objects[ $_SID_to_add ].membership += @( $script:_current_group_dn_inloop )
        #Certain groups can appear several times, it is on purpose to also give some stats while we print out the account
    }
    
    Write-Debug "-- _IncludeObjects"
}
 
# Function for primaryid membeship to manage the case the user has a primary group
# different than the default one
function _RGetMembershipRid(  ) {
    param (
        $_SID_TO_RID
    )
    Write-Debug "++ _RGetMembershipRid"
    Write-Debug "Loop depth: $script:_loop_depth"
    # Add users via PrimaryGroup RID
    #!!! ADD A CHECK IF IT IS A BUILT-IN GROUP
    $_current_group_member_primarygroupid_rid = $_SID_TO_RID.Split("-")[-1]
    # If not domain users, we list
    Switch ( $_current_group_member_primarygroupid_rid ) {
        # that's also the occasion to detect if they add Domain users or computers
        "513" { 
            #"# OMG DOMAIN USERS" #Well this is not a good sign
            Write-Debug "WARNING: Domain Users is included in this group."
            $script:_issues += @{
                "Title"    = "Well-known security principal is a member of a group" 
                "Objects"  = "The group Domain Users is a member of $script:_current_group_name"
                "Severity" = 3
            }
        }
        "515" {
            #"# OMG DOMAIN COMPUTERS" #rare but we never know...
            Write-Debug "WARNING: Domain Computers is included in this group."
            $script:_issues += @{
                "Title"    = "Well-known security principal is a member of a group" 
                "Objects"  = "The group Domain Computers is a member of $script:_current_group_name"
                "Severity" = 3
            }
        }
        default { # else let's try to get what group
            $script:_loop_depth++
            Write-Debug "Incrementing the loop depth to: $script:_loop_depth"
            Write-Debug "Instanciating a System.DirectoryServices.DirectorySearcher object" 
            $_current_group_member_primarygroupid = New-Object System.DirectoryServices.DirectorySearcher
            $_current_group_member_primarygroupid_filter = "(&(saMAccountName=*)(PrimaryGroupID=$_current_group_member_primarygroupid_rid))"
            $_current_group_member_primarygroupid.SearchRoot = "LDAP://$script:_current_dc_in_use"
            $_current_group_member_primarygroupid.Filter = $_current_group_member_primarygroupid_filter
            $_current_group_member_primarygroupid.SearchScope = "Subtree"
            $_current_group_member_primarygroupid.PageSize = 1000
            $_current_group_member_primarygroupid_result = $_current_group_member_primarygroupid.FindAll()
            Write-Debug "Call FindAll()"
            $_current_group_member_primarygroupid_result | ForEach-Object {
                # Get the SID of the group
                $_current_group_member_primarygroupid_result_sid = (New-Object Security.Principal.Securityidentifier(($_.GetDirectoryEntry()).objectSid[0],0)).Value
                # We got it! let's add it to the list...
                Switch ( ($_.GetDirectoryEntry()).objectCategory ) {
                    "CN=Person,$_forest_schema" {
                        $_ActualMember_item_type = "P" }
                    "CN=Computer,$_forest_schema" {
                        $_ActualMember_item_type = "C" }
                }
                $_ActualMember_item = @{
                    ActualMemberLevel = $script:_loop_depth
                    ActualMemberType = $_ActualMember_item_type
                    ActualMemberDn   = $( [string] ($_.GetDirectoryEntry()).distinguishedName)
                    ActualMemberPrimary = $true
                    ActualMemberSID = $_current_group_member_primarygroupid_result_sid
                    ActualMemberOcc = $_occurence_history
                }
                # We add the guy to the hashtable
                Write-Debug "Adding $( [string] ($_.GetDirectoryEntry()).distinguishedName) via PrimaryGroupID"
                $script:_ht_sid_domain[$script:_current_domain_sid].DomainGroups[$script:_current_group_name].ActualMembers += $_ActualMember_item
                _IncludeObjects $_current_group_member_primarygroupid_result_sid $_ActualMember_item_type
            } # End of the loop for the user having the primary ID set with the group RID
            $script:_loop_depth--
            Write-Debug "Decrementing the loop depth to: $script:_loop_depth"
        } # End of the default
    } # End of the Switch
    Write-Debug "-- _RGetMembershipRid"
} # End of the function _RGetMembershipRid
 
 
function _RGetMembership()
{
    param (
        [string] $_SID 
    )
 
    # Tracking IN/OUT of the recursive function
    Write-Debug "++ _RGetMembership"
    Write-Debug "Loop depth: $script:_loop_depth"
 
 
    # Get the membership of the group
    $_current_group_members = ([ADSI]"LDAP://$script:_current_dc_in_use/<SID=$_SID>").member
    $_current_group_members | ForEach-Object `
    {
        #! is that line an issue?
        $_current_group_member = [ADSI]"GC://$_"
        $_current_group_member_dn = $($_current_group_member.distinguishedName)
        
        # If we can get a SID, we can get a target DC and start the expension
        If ( $_current_group_member.objectSid )
        {
            $_current_group_member_sid = (New-Object Security.Principal.Securityidentifier(($_current_group_member).objectSid[0],0)).Value
            Write-Debug "Member SID: $_current_group_member_sid"
            # Add the discovered SID to the history occurence table
            $script:_history_sid += $_current_group_member_sid
            # Count the number of occurence of the same SID
            $_occurence_history = ($script:_history_sid | Group-Object -NoElement | Where-Object { $_.Name -eq $_current_group_member_sid }).Count
            Write-Debug "Member occurence: $_occurence_history"
            
            # Depending of the type of object, we do something
            Write-Debug "Entering the Switch statement to detect the object category"
            Switch ( $_current_group_member.objectCategory)
            {
                "CN=Group,$_forest_schema"
                {
                    Write-Debug "The member is a Group"
                    # Check if the group already showed up on the history
                    If ( $script:_loop_sid -notcontains $_current_group_member_sid )
                    {
                        $script:_loop_sid += $_current_group_member_sid
                        Write-Debug "First time we encounter this group for this enumeration"
                        
                        # We make sure we target the best DC for the group enumeration
                        $_dc_to_target = $script:_ht_sid_domain.Item( $_current_group_member_sid.Split("-")[0..6] -Join "-" ).DomainDC
                        If ( $_dc_to_target -eq $null )
                        {
                            Write-Debug "Cannot find a suitable DC. Carry on with: $script:_current_dc_in_use"
                        } Else {
                            $script:_current_dc_in_use = $_dc_to_target
                        }
                        Write-Debug "Set DC to use to: $script:_current_dc_in_use"

                        Write-Debug   "Calling System.DirectoryServices.DirectorySearcher for $_current_group_dn"
                        $_GetGroupMetadata = New-Object System.DirectoryServices.DirectorySearcher( [ADSI]"LDAP://$($script:_current_dc_in_use)/$($_current_group_member_dn)", "(objectClass=*)" , "msDS-ReplValueMetaData" )
                        Write-Debug   "Store the msDS-ReplValueMetaData of the group"
                        Try   { $_current_groupmetadata = $_GetGroupMetadata.FindOne().Properties."msds-replvaluemetadata" }
                        Catch { Write-Debug "Cannot get the msDS-ReplValueMetaData for $_current_group_member_dn" ; $_current_groupmetadata = $null }

                        # We add the member to the $_ht_sid_domain
                        $_member_item = @{
                            ActualMemberType  = "G"
                            ActualMemberLevel = $script:_loop_depth
                            ActualMemberDn    = $_current_group_member_dn
                            ActualMemberOcc   = $_occurence_history
                            ActualMemberSID   = $_current_group_member_sid
                            GroupMetadata     = $_current_groupmetadata
                        }
                        $script:_ht_sid_domain[$script:_current_domain_sid].DomainGroups[$script:_current_group_name].ActualMembers += $_member_item
                        # Add the group to the ht_sid_objects
                        # Well no, why would I do that? Ok I do it, we'll see
                        #_IncludeObjects $_current_group_member_sid "G"

                        $script:_current_group_dn_inloop += @( $_current_group_member_dn )

                        # Include RID member
                        Write-Debug "Calling _RGetMembershipRid with $_current_group_member_sid"
                        _RGetMembershipRid $_current_group_member_sid
                        # We loop to the same function
                        $script:_loop_depth++
                        Write-Debug "Incrementing the loop depth to: $script:_loop_depth"
                        Write-Debug "Calling _RGetMembership with $_current_group_member_sid"
                        _RGetMembership $_current_group_member_sid
                        $script:_loop_depth--
                        Write-Debug "Decrementing the loop depth to: $script:_loop_depth"
 
                        
                    } Else {
                        Write-Debug   "Already encountered this group for this enumeration"
                        Write-Verbose "Group already nested. This is a duplicate."
                        $script:_issues += @{
                            "Title"    = "Loop detected in group membership. The group $_current_group_member_dn is already in $script:_current_group_name."
                            "Severity" = 1
                        }
                    }
 
                }
                "CN=Person,$_forest_schema"
                {
                    Write-Debug "The member is a Person"
                    # Add the membership
                    $_member_item = @{
                        ActualMemberType  = "P"
                        ActualMemberLevel = $script:_loop_depth
                        ActualMemberDn    = $_current_group_member_dn
                        ActualMemberOcc   = $_occurence_history
                        ActualMemberSID   = $_current_group_member_sid
                    }
                    $script:_ht_sid_domain[$script:_current_domain_sid].DomainGroups[$script:_current_group_name].ActualMembers += $_member_item
                    # Add the user to the ht_sid_objects
                    _IncludeObjects $_current_group_member_sid "P"
                }
                "CN=Computer,$_forest_schema"
                {
                    Write-Debug "The member is a Computer"
                    # Add the membership
                    $_member_item = @{
                        ActualMemberType  = "P"
                        ActualMemberLevel = $script:_loop_depth
                        ActualMemberDn    = $_current_group_member_dn
                        ActualMemberOcc   = $_occurence_history
                        ActualMemberSID   = $_current_group_member_sid
                    }
                    $script:_ht_sid_domain[$script:_current_domain_sid].DomainGroups[$script:_current_group_name].ActualMembers += $_member_item
                    # Add the user to the ht_sid_objects
                    _IncludeObjects $_current_group_member_sid "C"
                }
                "CN=Foreign-Security-Principal,$_forest_schema"
                {
                    # Check if it is a well known foreign-scurity-principal
                    If ( $_wellKnownSecurity.Keys -contains $_current_group_member_sid )
                    {
                        Write-Verbose "Well-known security Foreign-Security-Principal included: $($_wellKnownSecurity[$_current_group_member_sid])"
                        $script:_issues += @{
                            "Title"    = "The wellknown security principal $($_wellKnownSecurity[$_current_group_member_sid]) is a member of $script:_current_group_name."
                            "Severity" = 1
                        }
 
                    }
                    Write-Debug "The member is a Foreign-Security-Principal"
                    # Add the membership
                    $_member_item = @{
                        ActualMemberType  = "F"
                        ActualMemberLevel = $script:_loop_depth
                        ActualMemberDn    = $_current_group_member_dn
                        ActualMemberOcc   = $_occurence_history
                        ActualMemberSID   = $_current_group_member_sid
                    }
                    $script:_ht_sid_domain[$script:_current_domain_sid].DomainGroups[$script:_current_group_name].ActualMembers += $_member_item
                    # Add the user to the ht_sid_objects
                    _IncludeObjects $_current_group_member_sid "F"
                }
            }
        } Else {
            Write-Debug "We cannot get the objectSid for $_current_group_member_dn"
        }
    }
    # Tracking IN/OUT of the recursive function
    Write-Debug "-- _RGetMembership"
}
 
Write-Debug "Enumerating all domains and their groups"
# Enumerating all domains and get the memberships of all designated groups
$_i_domains = 0
# This cannot be included in the previous domain enumeration because we potentially need info from all domain to expend the membership
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object `
{
    $_current_domain_sid  = $_.name
    $_current_domain_dns  = $_.value.DomainName
    # Display a progress bar 
    Write-Progress -Activity "Gathering group membership.." -Status "$($_i_domains+1) / $($_forest.Domains.Count)" -CurrentOperation "Collecting group membership for the domain $_current_domain_dns.." -Id 0 -PercentComplete ( $_i_domains++ / $_forest.Domains.Count * 100 )
    Write-Verbose "Domain: $_current_domain_dns"
    $script:_current_dc_in_use = $_.value.DomainDC
    # Store the current DC to reinitialize the DC when enumerating groups
    $_dc_to_use = $script:_current_dc_in_use
    Write-Debug "DC in use: $script:_current_dc_in_use"
    Write-Debug "Enumerating groups..."
    # Enumerate all groups we want to get the membership on this domain
    $_i_groups = 0
    $_total_groups = $_.value.DomainGroups.Count
    $_.value.DomainGroups.GetEnumerator() | ForEach-Object `
    {
        $_current_group_sid  = $($_.value.SID)
        $script:_current_group_AdminSDHolderProtected = $($_.value.AdminSDHolderProtected)
        $script:_current_group_Admins = $($_.value.Admins)
        $script:_current_group_name = $($_.name)
        # Display a progress bar 
        Write-Progress -Activity "Collecting members.." -Status "$($_i_groups+1) / $($_total_groups)" -CurrentOperation "Group: $script:_current_group_name.." -Id 1 -PercentComplete ( $_i_groups++ / $_total_groups * 100 ) -ParentId 0
        Write-Debug "Group SID: $_current_group_sid"
        Write-Debug "Group generic name: $script:_current_group_name"
        $_current_group_actual_name = $(([ADSI]"LDAP://$script:_current_dc_in_use/<SID=$_current_group_sid>").name)
        $_current_group_dn = $(([ADSI]"LDAP://$script:_current_dc_in_use/<SID=$_current_group_sid>").distinguishedName)
        $script:_ht_sid_domain[$_current_domain_sid].DomainGroups[$script:_current_group_name].ActualName = $_current_group_actual_name
        $script:_ht_sid_domain[$_current_domain_sid].DomainGroups[$script:_current_group_name].ActualDN = $_current_group_dn
        Write-Verbose "Group actual name: $_current_group_actual_name"
        Write-Debug   "Group DN: $_current_group_dn"
        $script:_current_group_dn_inloop = @( $_current_group_dn ) # used to trace membership per user
        Write-Verbose "Get membership..."
        # Reseting the $script:_history_sid containing every occurences of objects
        $script:_history_sid = @()
        $script:_loop_sid    = @()
        # We keep track of the depth of loop that we are making to arrange graphical output later
        $script:_loop_depth  = 0
 
        Write-Debug   "Calling _RGetMembership for the first time on that loop with $_current_group_sid"
        _RGetMembership $_current_group_sid
 
        Write-Debug   "Calling _RGetMembershipRid for the first time on that loop with $_current_group_sid"
        _RGetMembershipRid $_current_group_sid # This call will not return anything anyways for built-in groups

        Write-Debug   "Calling System.DirectoryServices.DirectorySearcher for $_current_group_dn"
        $_GetGroupMetadata = New-Object System.DirectoryServices.DirectorySearcher( [ADSI]"LDAP://$($script:_current_dc_in_use)/$($_current_group_dn)", "(objectClass=*)" , "msDS-ReplValueMetaData" )
        Write-Debug   "Store the msDS-ReplValueMetaData of the group"
        $script:_ht_sid_domain[$_current_domain_sid].DomainGroups[$script:_current_group_name].GroupMetadata = $_GetGroupMetadata.FindOne().Properties."msds-replvaluemetadata"
        # Reinit the DC to the current DC for the domain
        $script:_current_dc_in_use = $_dc_to_use
    }
    Write-Verbose "End of $script:_current_group_name enumeration for the domain $_current_domain_dns"
}
Write-Verbose "End of all groups enumeration"
 
# Get object's information
$_i_objects = 0 # used for the Write-Progress bar
Write-Verbose "Collection object's information"
$script:_ht_sid_objects.GetEnumerator() | ForEach-Object{ # For each of the objects we previsouly colected
 
    Write-Progress -Activity "Gathering object's information.." -Status "$($_i_objects+1) / $($_ht_sid_objects.Count)" -CurrentOperation "Collecting information for $($_.Name).." -Id 0 -PercentComplete ( $_i_objects++ / $_ht_sid_objects.Count * 100 )
    #Check if it is a Foreign-
    $_object_type = $($_.Value.type)
    $_object_sid  = $($_.Name)
    Write-Debug "Object type: $_object_type"
    Write-Debug "Object SID: $_object_sid"
    Switch ( $_object_type )
    {
        "F"
        {
            # Do something this is a F
            Write-Debug "Foreign-Security-Principal"
            #If ( $_wellKnownSecurity.Keys -contains $_object_sid )
            #{
            #    #write something special
            #}
            #collect the info
            #Write-Debug "Using DC: $_dc_to_target"
            #$_object = [ADSI]"GC://<SID=$_object_sid>"            
            #$script:_ht_sid_objects[ $_object_sid ].distinguishedName = $($_object.distinguishedName)
            #$script:_ht_sid_objects[ $_object_sid ].objectCategory = $($_object.objectCategory)
            #$script:_ht_sid_objects[ $_object_sid ].displayname = $($_object.displayName)
        }
        "G"
        {
            #If it is a group, we collect group type and scope, plus date of creation and modification
            #!!! NEED TO BE IMPLEMENTED

        }
        Default # Which is P or C # Potential I later?
        {
            #decide what DC to use
            $_dc_to_target   = $script:_ht_sid_domain.Item( $_object_sid.Split("-")[0..6] -Join "-" ).DomainDC
            $_object_meta_dc = $script:_ht_sid_domain.Item( $_object_sid.Split("-")[0..6] -Join "-" ).DomainMeta
            If ( $_dc_to_target -eq $null )
            {
                Write-Debug "Cannot find a suitable DC. Attribute collection might fail."
            }
            #collect the info
            Write-Debug "Using DC: $_dc_to_target"
            $_object = [ADSI]"LDAP://$_dc_to_target/<SID=$_object_sid>"
            # Some conversion stuff with the ntte epoch of some attributes
            # We do try/catch to set the date to Never if there's no value or an invalid input
            # Colection the accountExpires attribute
            If ( $_object.ConvertLargeIntegerToInt64( $_object.accountExpires.value ) -ne 0 )
            {
                Try
                {
                    $_object_accountExpires = [datetime]::fromfiletime( $_object.ConvertLargeIntegerToInt64( $_object.accountExpires.value ))
                }
                Catch
                {
                    $_object_accountExpires = "Never"
                }
            } Else {
                $_object_accountExpires = "Never"
            }
            # Colection the lastLogonTimestamp attribute
            Try
            {
                $_object_lastLogonTimestamp = [datetime]::fromfiletime( $_object.ConvertLargeIntegerToInt64( $_object.lastLogonTimestamp.value )) 
                $_object_lastLogonTimestamp = $_object_lastLogonTimestamp.ToUniversalTime()
            }
            Catch
            {
                $_object_lastLogonTimestamp = "Never"
            }
            # Colection the pwdLastSet attribute
            Try
            {
                $_object_pwdLastSet = [datetime]::fromfiletime( $_object.ConvertLargeIntegerToInt64( $_object.pwdLastSet.value ))
                $_object_pwdLastSet = $_object_pwdLastSet.ToUniversalTime()
            
            }
            Catch
            {
                $_object_pwdLastSet = "Never"
            }
            # Colection the lockoutTime attribute
            Try
            {
                $_object_lockoutTime = [datetime]::fromfiletime( $_object.ConvertLargeIntegerToInt64( $_object.lockoutTime.value ))
                $_object_lockoutTime = $_object_lockoutTime.ToUniversalTime()
            }
            Catch
            {
                $_object_lockoutTime = ""
            }
            # For the sidHistory we just do a yes/no.. come on... For bow at lest
            If ( $_object.sidHistory )
            {
	            $_object_sidHistory = "Yes"
            } Else {
	            $_object_sidHistory = "No"
            }
            # Collecting replication metadata information
            Write-Debug "Calling GetReplicationMetadata(...)"
            $_object_meta = $_object_meta_dc.GetReplicationMetadata($($_object.distinguishedName))
            Try
            {
                $_object_password_version = $($_object_meta.unicodepwd.Version)
                # Get if the adminCount is set let's check when it has been done
                If ( $($_object.adminCount) -eq 1 )
                {
                    $_object_admincount_date = $( ($_object_meta.admincount.LastOriginatingChangeTime).ToUniversalTime() )
                } Else {
                    $_object_admincount_date = ""
                }
            }
            Catch
            {
                Write-Debug "something is wrong with the metadata collection of $($_object.distinguishedName) but we move on!"
            }
            If ( $($_object.userPrincipalName) -ne $null ) {
                Try
                {
                    $userPrincipalNameVersion = $($_object_meta.userprincipalname.Version)
                    $userPrincipalNameTime    = $(($_object_meta.userprincipalname.LastOriginatingChangeTime).ToUniversalTime())
                }
                Catch
                {
                    # Cannot get metainformation, set to blank and carry on
                    $userPrincipalNameVersion = ""
                    $userPrincipalNameTime    = ""
                }
            }
            # Prepare the principal info and put it in the hashtable
            $script:_ht_sid_objects[ $_object_sid ].distinguishedName = $($_object.distinguishedName)
            $script:_ht_sid_objects[ $_object_sid ].objectCategory = $($_object.objectCategory)
            $script:_ht_sid_objects[ $_object_sid ].sAMAccountName = $($_object.sAMAccountName)
            $script:_ht_sid_objects[ $_object_sid ].userPrincipalName = $($_object.userPrincipalName)
            $script:_ht_sid_objects[ $_object_sid ].userPrincipalNameVersion = $userPrincipalNameVersion
            $script:_ht_sid_objects[ $_object_sid ].userPrincipalNameTime = $userPrincipalNameTime
            $script:_ht_sid_objects[ $_object_sid ].userAccountControl = $($_object.userAccountControl)
            $script:_ht_sid_objects[ $_object_sid ].whenCreated = $($_object.whenCreated)
            $script:_ht_sid_objects[ $_object_sid ].lastLogonTimestamp = $_object_lastLogonTimestamp
            $script:_ht_sid_objects[ $_object_sid ].lockoutTime = $_object_lockoutTime
            $script:_ht_sid_objects[ $_object_sid ].pwdLastSet = $_object_pwdLastSet
            $script:_ht_sid_objects[ $_object_sid ].passwordVersion = $_object_password_version
            $script:_ht_sid_objects[ $_object_sid ].description = $($_object.description)
            $script:_ht_sid_objects[ $_object_sid ].mail = "$($_object.mail) "
            $script:_ht_sid_objects[ $_object_sid ].accountExpires = $_object_accountExpires
            $script:_ht_sid_objects[ $_object_sid ].adminCount = $($_object.adminCount)
            $script:_ht_sid_objects[ $_object_sid ].adminCountDate = $_object_admincount_date
            $script:_ht_sid_objects[ $_object_sid ].sidHistory = $_object_sidHistory
            $script:_ht_sid_objects[ $_object_sid ].userPassword = $($_object.userPassword)
            $script:_ht_sid_objects[ $_object_sid ].userWorkstations = $($_object.userWorkstations)
            $script:_ht_sid_objects[ $_object_sid ].displayname = $($_object.displayName)
 
            #!!! ADD SOME LIVE ANALYSIS to fire issue
        }
    }
    Write-Debug "> Next object"
}
Write-Verbose "Objects' attributes collection done"
 
Write-Verbose "Starting collection of objects with adminCount=1"
Write-Debug   "Enumerating all domains and their object with adminCount set to 1"
# Enumerating all domains and get the objects with admincount=1
$_i_domains = 0
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object `
{
    $_current_domain_sid  = $_.name
    $_current_domain_dns  = $_.value.DomainName
    Write-Debug "Domain: $_current_domain_dns"
    $_current_domain_dc   = $_.value.DomainDC
    $_current_domain_meta = $_.value.DomainMeta
 
    # Display a progress bar 
    Write-Progress -Activity "Enumerating objects.." -Status "$($_i_domains+1) / $($_forest.Domains.Count)" -CurrentOperation "Collecting object with the adminCount set to 1 in $_current_domain_dns.." -Id 0 -PercentComplete ( ++$_i_domains / $_forest.Domains.Count * 100 )
 
    # Detect orphan admincount
    Write-Debug "Instanciating a System.DirectoryServices.DirectorySearcher object" 
    $_admincount_object = New-Object System.DirectoryServices.DirectorySearcher
    $_admincount_object.SearchRoot = "LDAP://$_current_domain_dc"
    # We exclude the krbTgt account this one is protected and not a member of any protected group
    $_admincount_object.Filter = "(&(|(objectCategory=CN=Person,$_forest_schema)(objectCategory=CN=Computer,$_forest_schema))(adminCount=1)(!(samaccountname=krbtgt)))"
    $_admincount_object.SearchScope = "Subtree"
    $_admincount_object.PageSize = 1000
    Write-Debug "Calling FindAll()"
    $_admincount_objects = $_admincount_object.FindAll()
    $_i_admincount_objects = 0 # used for displaying the progress bar
    $_i_admincount_total   = $_admincount_objects.Count # used for displaying the progress bar
    Write-Debug "Found $_i_admincount_total to check"
    $_admincount_objects | ForEach-Object `
    {
        $_current_object_admincount = $( [string] ($_.GetDirectoryEntry()).distinguishedName)
        $_current_object_admincount_uac = $( [string] ($_.GetDirectoryEntry()).userAccountControl)
        $_current_object_admincount_cat = $( [string] ($_.GetDirectoryEntry()).objectCategory)
        #?Need to investigate why the Count can return $null and still have records in the RS
        Try   { Write-Progress -Activity "Listing account with adminCount=1.." -Status "$($_i_admincount_objects+1) / $([string] $_i_admincount_total)" -CurrentOperation "Checking object $_current_object_admincount.." -Id 1 -PercentComplete ( ++$_i_admincount_objects / $_i_admincount_total * 100 ) -ParentId 0 }
        Catch { Write-Debug "Error displaying the progress bar..." }
        If (($script:_ht_sid_objects.Values.GetEnumerator() | Where-Object { $_.AdminSDHolderProtected -eq 1 } | ForEach-Object { $_.distinguishedName }) -contains $_current_object_admincount )
        {
            Write-Debug "$_current_object_admincount is listed as protected"
        } Else {
            Write-Verbose "$_current_object_admincount is an orphan, getting info..."
            Write-Debug "Calling GetReplicationMetadata(...)"
            $_object_meta = $_current_domain_meta.GetReplicationMetadata($_current_object_admincount)
            Try   { $_object_admincount_att = $( ($_object_meta.admincount.LastOriginatingChangeTime).ToUniversalTime() ) }
            Catch { $_object_admincount_att = "" }
            Try   { $_object_admincount_sec = $( ($_object_meta.ntsecuritydescriptor.LastOriginatingChangeTime).ToUniversalTime() ) }
            Catch { $_object_admincount_sec = "" }
            $_ht_admincount += @{ $_current_object_admincount = @{ "userAccountControl" = $_current_object_admincount_uac; "objectCategory" = $_current_object_admincount_cat ; "adminCountWhenModified" = $_object_admincount_att ; "ntSecurityDescriptorWhenModified" = $_object_admincount_sec} } 
            $_issues += @{ "Title" = "Orphaned adminSHolder objects found" ; "Objects" = "($_current_object_admincount) objects have been detected" ; "Severity" = 1 }
        }
    }
}
Write-Verbose "End of collection of objects with adminCount=1"
 
# Counting unique admins per domain an0d membership statistics
# YES this part deserve more comments...
$_i_domains = 0
$_unique_admins_forest = @()
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object `
{
    $_unique_admins_domain = @()
    $_current_domain_sid  = $_.name
    $_current_domain_dns  = $_.value.DomainName
    Write-Debug "Domain: $_current_domain_dns"
    # Display a progress bar 
    Write-Progress -Activity "Counting unique administrative accounts.." -Status "$($_i_domains+1) / $($_forest.Domains.Count)" -CurrentOperation "For the domain $_current_domain_dns.." -Id 0 -PercentComplete ( ++$_i_domains / $_forest.Domains.Count * 100 )
    # For each domain, list all groups and get membership for groups where Admins = 1
    $_.value.DomainGroups.GetEnumerator() | ForEach-Object `
    {
        $_current_group_enum = $_.Name
        $_stats_direct_members_total = 0
        $_stats_direct_members_c     = 0
        $_stats_direct_members_f     = 0
        $_stats_direct_members_p     = 0
        $_stats_direct_members_g     = 0
        $_stats_nested_members_total = 0
        $_stats_nested_members_c     = 0
        $_stats_nested_members_f     = 0
        $_stats_nested_members_p     = 0
        $_stats_nested_members_g     = 0
        $_track_unique_dn_in_group   = @()
 
        $script:_ht_sid_domain[ $_current_domain_sid].DomainGroups[$_current_group_enum].ActualMembers.GetEnumerator() | ForEach-Object {
            $_current_group_enum_type  = $_.ActualMemberType 
            $_current_group_enum_dn    = $_.ActualMemberDn
            $_current_group_enum_level = $_.ActualMemberLevel
            # !! Track for uniqueness
            If ( $_track_unique_dn_in_group -contains $_current_group_enum_dn ) {
                Write-Debug "Member $_current_group_enum_dn is already in the list, skip it for the stats"
            } Else {
                $_track_unique_dn_in_group += $_current_group_enum_dn
 
                If ( $_current_group_enum_level -eq 0 ) {
                    $_stats_direct_members_total++
                    Switch ( $_current_group_enum_type )
                    {
                        "C" { $_stats_direct_members_c++ }
                        "F" { $_stats_direct_members_f++ }
                        "G" { $_stats_direct_members_g++ }
                        "P" { $_stats_direct_members_p++ }
                    }
                } Else {
                    $_stats_nested_members_total++
                    Switch ( $_current_group_enum_type )
                    {
                        "C" { $_stats_nested_members_c++ }
                        "F" { $_stats_nested_members_f++ }
                        "G" { $_stats_nested_members_g++ }
                        "P" { $_stats_nested_members_p++ }
                    }
 
                }
                # We also exclude the group from the stats
                If ( $script:_ht_sid_domain[ $_current_domain_sid].DomainGroups[$_current_group_enum].Admins -eq $true -and $_unique_admins_domain -notcontains $_current_group_enum_dn -and $_current_group_enum_type -ne "G" )
                {
                    #Admins Domain +1
                    $_unique_admins_domain += $_current_group_enum_dn
                    If ( $_unique_admins_forest -notcontains $_current_group_enum_dn )
                    {
                         #Admins Forest +1
                         $_unique_admins_forest += $_current_group_enum_dn
                    }
                }
            
            }
        }
        $script:_ht_sid_domain[ $_current_domain_sid ].DomainGroups[$_current_group_enum ].Stats = @{
                "DirectMembersTotal" = $_stats_direct_members_total
                "DirectMembersTotalNoGroups" = $_stats_direct_members_total - $_stats_direct_members_g
                "DirectMembersC"     = $_stats_direct_members_c
                "DirectMembersF"     = $_stats_direct_members_f
                "DirectMembersP"     = $_stats_direct_members_p
                "DirectMembersG"     = $_stats_direct_members_g
                "NestedMembersTotal" = $_stats_nested_members_total + $_stats_direct_members_total
                "NestedMembersTotalNoGroups" = $_stats_nested_members_total + $_stats_direct_members_total - $_stats_nested_members_g - $_stats_direct_members_g
                "NestedMembersC"     = $_stats_nested_members_c     + $_stats_direct_members_c
                "NestedMembersF"     = $_stats_nested_members_f     + $_stats_direct_members_f
                "NestedMembersP"     = $_stats_nested_members_p     + $_stats_direct_members_p
                "NestedMembersG"     = $_stats_nested_members_g     + $_stats_direct_members_g
        }
    }
    # Store the nb of unique admin accounts in the domain
    $script:_ht_sid_domain[ $_current_domain_sid ].TotalUniqueAdmins = $_unique_admins_domain.Count
}
# Store the nb of unique admin accounts within the forest
$script:_ht_forest.TotalUniqueAdmins = $_unique_admins_forest.Count
 
Write-Verbose "Starting collection of objects with a tricky userAccountName values"
Write-Debug   "Enumerating all domains and their object with a tricky userAccountName values"
<# What is a tricky userAccountName value?
    32 0x20 = "PASSWD_NOTREQD" ,
    128 0x80 = "ENCRYPTED_TEXT_PWD_ALLOWED" ,
    65536 0x10000 = "DONT_EXPIRE_PASSWORD" ,
    524288 0x80000 = "TRUSTED_FOR_DELEGATION" ,
    2097152 0x200000 = "USE_DES_KEY_ONLY" ,
    4194304 0x400000 = "DONT_REQ_PREAUTH" ,
    16777216 0x1000000 = "TRUSTED_TO_AUTH_FOR_DELEGATION" ,
    33554432 0x2000000 = "NO_AUTH_DATA_REQUIRED",
#>
# Translation table for the userAccountControl attribute
$_UACPattern = @{
    0x1 = "SCRIPT" ,
    "The logon script will be run." ; 
    0x2 = "ACCOUNTDISABLE" ,
    "The user account is disabled. " ;
    0x8 = "HOMEDIR_REQUIRED",
    "The home folder is required." ;
    0x10 = "LOCKOUT",
    "The account is locked-out." ;
    0x20 = "PASSWD_NOTREQD" ,
    "No password is required." ;
    0x40 = "PASSWD_CANT_CHANGE",
    "The user cannot change the password. This is a permission on the user's object." ;
    0x80 = "ENCRYPTED_TEXT_PWD_ALLOWED" ,
    "The user can send an encrypted password." ;
    0x100 = "TEMP_DUPLICATE_ACCOUNT",
    "This is an account for users whose primary account is in another domain.";
    0x200 = "NORMAL_ACCOUNT",
    "This is a default account type that represents a typical user." ;
    0x800 = "INTERDOMAIN_TRUST_ACCOUNT", 
    "This is a permit to trust an account for a system domain that trusts other domains." ;
    0x1000 = "WORKSTATION_TRUST_ACCOUNT" ,
    "This is a computer account for a computer that is running Microsoft Windows and is a member of this domain." ;
    0x2000 = "SERVER_TRUST_ACCOUNT" ,
    "This is a computer account for a domain controller that is a member of this domain." ;
    0x10000 = "DONT_EXPIRE_PASSWORD" ,
    "Represents the password, which should never expire on the account." ;
    0x20000 = "MNS_LOGON_ACCOUNT" ,
    "This is an MNS logon account." ;
    0x40000 = "SMARTCARD_REQUIRED" ,
    "When this flag is set, it forces the user to log on by using a smart card."
    0x80000 = "TRUSTED_FOR_DELEGATION" ,
    "When this flag is set, the service account under which a service runs is trusted for Kerberos delegation." ;
    0x100000 = "NOT_DELEGATED" ,
    "When this flag is set, the security context of the user is not delegated to a service even if the service account is set as trusted for Kerberos delegation." ;
    0x200000 = "USE_DES_KEY_ONLY" ,
    "Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys." ;
    0x400000 = "DONT_REQ_PREAUTH" ,
    "This account does not require Kerberos pre-authentication for logging on." ;
    0x800000 = "PASSWORD_EXPIRED" ,
    "The user's password has expired." ;
    0x1000000 = "TRUSTED_TO_AUTH_FOR_DELEGATION" ,
    "The account is enabled for delegation. This is a security-sensitive setting.." ;
    0x2000000 = "NO_AUTH_DATA_REQUIRED",
    "This bit indicates that when the Key Distribution Center (KDC) is issuing a service ticket for this account, the Privilege Attribute Certificate (PAC) MUST NOT be included." ;
    0x4000000 = "PARTIAL_SECRETS_ACCOUNT" ,
    "The account is a read-only domain controller (RODC)." ;
}
$_i_domains = 0
$_ht_tricky_uac = @{"PASSWD_NOTREQD" = @();"ENCRYPTED_TEXT_PWD_ALLOWED"=@();"DONT_EXPIRE_PASSWORD"=@();"TRUSTED_FOR_DELEGATION"=@();"USE_DES_KEY_ONLY"=@();"DONT_REQ_PREAUTH"=@();"TRUSTED_TO_AUTH_FOR_DELEGATION"=@();"NO_AUTH_DATA_REQUIRED"=@()}
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object `
{
    $_current_domain_sid  = $_.name
    $_current_domain_dns  = $_.value.DomainName
    Write-Debug "Domain: $_current_domain_dns"
    $_current_domain_dc   = $_.value.DomainDC
    $_current_domain_meta = $_.value.DomainMeta
 
    # Display a progress bar 
    Write-Progress -Activity "Gathering user objects with tricky UAC.." -Status "$($_i_domains+1) of $($_forest.Domains.Count)" -CurrentOperation "Collecting group membership for the domain $_current_domain_dns.." -Id 0 -PercentComplete ( ++$_i_domains / $_forest.Domains.Count * 100 )
 
    # Detect orphan admincount
    Write-Debug "Instanciating a System.DirectoryServices.DirectorySearcher object" 
    $_uac_object = New-Object System.DirectoryServices.DirectorySearcher
    $_uac_object.SearchRoot = "LDAP://$_current_domain_dc"
    #! Exlude DCs from the filter
    If ( $WideTrickyUAC -eq $true )
    {
        $_uac_filter = "(&(saMAccountName=*)(|(userAccountControl:1.2.840.113556.1.4.803:=128)(userAccountControl:1.2.840.113556.1.4.803:=524288)(userAccountControl:1.2.840.113556.1.4.803:=2097152)(userAccountControl:1.2.840.113556.1.4.803:=4194304)(userAccountControl:1.2.840.113556.1.4.803:=16777216)(userAccountControl:1.2.840.113556.1.4.803:=33554432)))"
    } Else {
        # We are probably in an environment with many accounts with a password not set or never expire
        $_uac_filter = "(&(saMAccountName=*)(|(userAccountControl:1.2.840.113556.1.4.803:=32)(userAccountControl:1.2.840.113556.1.4.803:=128)(userAccountControl:1.2.840.113556.1.4.803:=65536)(userAccountControl:1.2.840.113556.1.4.803:=524288)(userAccountControl:1.2.840.113556.1.4.803:=2097152)(userAccountControl:1.2.840.113556.1.4.803:=4194304)(userAccountControl:1.2.840.113556.1.4.803:=16777216)(userAccountControl:1.2.840.113556.1.4.803:=33554432)))"
    }
    $_uac_object.Filter = $_uac_filter
    $_uac_object.SearchScope = "Subtree"
    $_uac_object.PageSize = 1000
    Write-Debug "Calling FindAll() using filter: $_uac_filter"
    $_uac_objects = $_uac_object.FindAll()
    $_i_uac_objects = 0 # used for displaying the progress bar
    $_i_uac_total   = $_uac_objects.Count # used for displaying the progress bar
    Write-Debug "Found $_i_uac_total to check"
    $_uac_objects | ForEach-Object {
        $_current_object_uac_dn  = $( [string] ($_.GetDirectoryEntry()).distinguishedName)
        $_current_object_uac_uac = $( [string] ($_.GetDirectoryEntry()).userAccountControl)
        Write-Progress -Activity "Listing account with a tricky UAC value.." -Status "$_i_uac_objects of $_i_uac_total" -CurrentOperation "Checking object $_current_object_uac_dn.." -Id 1 -PercentComplete ( ++$_i_uac_objects / $_i_uac_total * 100 ) -ParentId 0
        #? Replace with a switch
        If ( $_current_object_uac_uac -band 0x20 )      { $_ht_tricky_uac["PASSWD_NOTREQD"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x80 )      { $_ht_tricky_uac["ENCRYPTED_TEXT_PWD_ALLOWED"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x10000 )   { $_ht_tricky_uac["DONT_EXPIRE_PASSWORD"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x80000 )   { $_ht_tricky_uac["TRUSTED_FOR_DELEGATION"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x200000 )  { $_ht_tricky_uac["USE_DES_KEY_ONLY"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x400000 )  { $_ht_tricky_uac["DONT_REQ_PREAUTH"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x1000000 ) { $_ht_tricky_uac["TRUSTED_TO_AUTH_FOR_DELEGATION"] += $_current_object_uac_dn }
        If ( $_current_object_uac_uac -band 0x2000000 ) { $_ht_tricky_uac["NO_AUTH_DATA_REQUIRED"] += $_current_object_uac_dn }
    }
}
Write-Verbose "End of collection of objects with tricky userAccountName values"
# Generate issues based on previous findings
$_ht_tricky_uac.GetEnumerator() | ForEach-Object `
{
    $_tricky_uac = $_.Name
    $_tricky_nb  = $_.Value.Count
    If ( $_tricky_nb -ne 0 )
    {
        $_issues += @{ "Title" = "Objects found with the flag $_tricky_uac" ; "Objects" = "($_tricky_nb objects) founds with the flag $_tricky_uac" ; "Severity" = 1 }
    }
}
# Calculate some stats about the collection
$_ScriptEndCollection  = [System.DateTime]::Now
$_ScriptCollectionTime = $_ScriptEndCollection - $_ScriptStartDate 
Write-Verbose "Collection time: $_ScriptCollectionTime"
 
 
$_ScriptAccountAnalisysStart  = [System.DateTime]::Now
# Some analysis of the user accounts
# Check for mail
$_check_mail = @()
# Check for password didnt change for a while
$_check_password_age = @()
# Check for password never changed since creation time (1 sec tolerence)
$_check_password_never = @()
# Check for account with a password which does not expire
$_check_password_noexpire = @()
# Account who didn't logout for a while or never
$_check_stale = @()
# Account with one more 1 UPN modification
$_check_upn = @()
 
# Now list all accounts
# Get object's information
$_i_objects = 0 # used for the Write-Progress bar
Write-Verbose "Account Analysis"
$_account_stale_threshold    = 180
$_account_password_threshold = 365
Write-Debug   "Threshold for inactive account: $_account_stale_threshold"
Write-Debug   "Threshold for password age: $_account_password_threshold"
$script:_ht_sid_objects.GetEnumerator() | Where-Object { $_.Value.type -ne "F" } | ForEach-Object `
{ # For each of the objects we previsouly colected
    $_current_object_name = $_.Value.displayName
    $_current_object_sid  = $_.Name
    $_current_object_cat  = $_.Value.objectCategory
    $_current_object_dn   = $_.Value.distinguishedName
    $_current_object_uac  = $_.Value.userAccountControl
    Write-Progress -Activity "Analysing objects.." -Status "$($_i_objects+1) / $($_ht_sid_objects.Count)" -CurrentOperation "Collecting information for $_current_object_name.." -Id 0 -PercentComplete ( ++$_i_objects / $_ht_sid_objects.Count * 100 )
    Write-Debug "Analyzing account $_current_object_sid"
    # Check if the account have an email
    If ( $_.Value.mail -like "*@*" -and $_.Value.Admin -eq $true)
    {
        #Add account to the mail list
        $_check_mail += @{ "distinguishedName" = $_current_object_dn ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "mail" = $_.Value.mail }
        $_issues += @{ "Title" = "Administrative account with an email address" ; "Objects" = "The account $_current_object_dn has a email address set ($($_.Value.mail))" ; "Severity" = 1 }
    }
 
    # Check if the password has changed since the limit
    $_inactive_t = ([System.DateTime]::Now).AddDays( -$_account_password_threshold )
    If ( $_.Value.whenCreated -le $_inactive_t -and $_.Value.Admin -eq $true -and $_.Value.pwdLastSet -le $_inactive_t )
    {
        #Add account to the password list
        $_check_password_age_delta = $( $(([System.DateTime]::Now) - $($_.Value.pwdLastSet)).Days )
        $_check_password_age += @{ "distinguishedName" = $_current_object_dn ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "pwdLastSet" = $_.Value.pwdLastSet ; "Delta" = $_check_password_age_delta }
        $_issues += @{ "Title" = "Admin account with an old password" ; "Objects" = "$_current_object_dn didn't change its password recently (pwdLastSet $_check_password_age_delta days ago)" ; "Severity" = 3 }
    }
                                                                                                                  
    #Check if the password never changed since creation
    $_logic_password_date_check = ( $_.Value.whenCreated -eq $_.Value.pwdLastSet ) -or (  ( $_.Value.whenCreated -le ($_.Value.pwdLastSet).AddSeconds(2) ) -and ( $_.Value.whenCreated -ge ($_.Value.pwdLastSet).AddSeconds(-2) ) ) 
    If ( $_.Value.Admin -eq $true -and  $_logic_password_date_check )
    {
        #Add account to the password didnt change since creation
        $_check_password_never += @{ "distinguishedName" = $_current_object_dn ; "whenCreated" = $_.Value.whenCreated ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "pwdLastSet" = $_.Value.pwdLastSet }
        $_issues += @{ "Title" = "Admin acount with a password which never changed" ; "Objects" = "$_current_object_dn never changed its password (whenCreated= $($_.Value.whenCreated))" ; "Severity" = 3 }
    }
 
    #Check if the password is set at no expire
    If ( $_.Value.Admin -eq $true -and $_.Value.userAccountControl -band 0x10000 )
    {
        #Add account to the password never expirelist
        $_check_password_noexpire += @{ "distinguishedName" = $_current_object_dn ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "pwdLastSet" = $_.Value.pwdLastSet }
        $_issues += @{ "Title" = "Admin account with a password which never expires" ; "Objects" = "$_current_object_dn has a password which does not expire" ; "Severity" = 2 }
    }
    # Check for stale admin account
    $_inactive_t = ([System.DateTime]::Now).AddDays( -$_account_stale_threshold )
    If ( $_.Value.whenCreated -le $_inactive_t -and $_.Value.Admin -eq $true -and ($_.Value.lastLogonTimestamp -le $_inactive_t -or $_.Value.lastLogonTimestamp -eq "Never" ) )
    {
        #Add account to the stale list
        If ( $_.Value.lastLogonTimestamp -eq "Never" )
        {
            $_check_stale_delta = "N/A"
        } Else {
            $_check_stale_delta = $( $(([System.DateTime]::Now) - $($_.Value.lastLogonTimestamp)).Days )
        }
        $_check_stale += @{ "distinguishedName" = $_current_object_dn ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "lastLogonTimestamp" = $_.Value.lastLogonTimestamp ; "Delta" = $_check_stale_delta }
        $_issues += @{ "Title" = "Stale admin account" ; "Objects" = "$_current_object_dn seems stale (lastLogonTimestamp = $($_.Value.lastLogonTimestamp) - $_check_stale_delta days ago)" ; "Severity" = 1 }
    }
    
    # Check for suspicious UPN modification
    If ( $_.Value.Admin -eq $true -and $_.Value.userPrincipalNameVersion -gt 1 )
    {
        #Add account to the UPN list
        $_check_upn += @{ "distinguishedName" = $_current_object_dn ; "userAccountControl" = $_current_object_uac ; "objectCategory" = $_current_object_cat ; "userPrincipalNameVersion" = $_.Value.userPrincipalNameVersion ; "userPrincipalNameTime" = $_.Value.userPrincipalNameTime ; "userPrincipalName" = $_.Value.userPrincipalName }
        $_issues += @{ "Title" = "Admin account with a suspicious UPN modification" ; "Objects" = "The account $_current_object_dn had several UPN modifications ($($_.Value.userPrincipalNameVersion) modifications)" ; "Severity" = 1 }
    }
}
 
$_ht_check_account = @{
    "MAIL" = $_check_mail
    "PASSWORD_AGE" = $_check_password_age
    "PASSWORD_NEVER" = $_check_password_never
    "PASSWORD_NOEXPIRE" = $_check_password_noexpire
    "STALE" = $_check_stale
    "UPN" = $_check_upn
}
$_ScriptAccountAnalisysEnd  = [System.DateTime]::Now
$_ScriptAccountAnalisysTime = $_ScriptAccountAnalisysEnd - $_ScriptAccountAnalisysStart 
Write-Verbose "Analysis time: $_ScriptAccountAnalisysTime"
 
 
# We check if the operator wants to skip the export to Clixml
If ( $SkipExport -eq $false )
{
    # Remove meta avant export otherwise it forces to refresh all properties, and that can take a while
    $_ht_sid_domain.GetEnumerator() | ForEach-Object `
    {
        $_.Value.DomainMeta = "N/A for export"
    }
 
    # Exporting to files for offline re-use
    $_ScriptStartExport  = [System.DateTime]::Now
    Write-Verbose "Exporting the data to XML files for offline use"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_forest.xml"
    $_ht_forest      | Export-Clixml "$($_collection_export_prefix)_ht_forest.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_sid_domain.xml"
    $_ht_sid_domain  | Export-Clixml "$($_collection_export_prefix)_ht_sid_domain.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_sid_objects.xml"
    $_ht_sid_objects | Export-Clixml "$($_collection_export_prefix)_ht_sid_objects.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_issue.xml"
    $_issues         | Export-Clixml "$($_collection_export_prefix)_issues.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_tricky_uac.xml"
    $_ht_tricky_uac  | Export-Clixml "$($_collection_export_prefix)_ht_tricky_uac.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_check_account.xml"
    $_ht_check_account  | Export-Clixml "$($_collection_export_prefix)_ht_check_account.xml"
    Write-Debug   "Generating $($_collection_export_prefix)_ht_admincount.xml"
    $_ht_admincount  | Export-Clixml "$($_collection_export_prefix)_ht_admincount.xml"
    $_ScriptEndExport  = [System.DateTime]::Now
    $_ScriptExportTime = $_ScriptEndExport - $_ScriptStartExport 
    Write-Verbose "Export time: $_ScriptExportTime"
} # End of SkipExport check
 
### GENERATE HTML
$_ScriptHTLMStart = [System.DateTime]::Now
Write-Verbose "Generating HTML output"
# Creating base64 image for the template to remove dependencies with PNG files in the same folder
$_ht_img = @{}
$_ht_img["forest"]     = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABYAAAAZCAYAAAA14t7uAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAADrSURBVEhL7ZVBCsIwEEV7FS/gUhAP6c6FegPBpRs3giBeQEGliHoAabVWaszoD1QnNE0txpV9MGTSTl5DGRJPCEG/CE9KSS6ARxPjYa29+jo4+IhTMc8r8Su3ips9n8b+iRIh1Yi5qQ7BPcAqnu4iVZQyP1yMdQhOYVdEyV0VpmDnpjoEp1Bcdsc8t4pbfZ8m27MqxIi5qQ4BeG4VlwnuAUZxvbOhwSKg8Kr/42MsaLgMqNHN7hzw3CgerUP1Mo/ZPsqsATw3iuOb/cR7NkdmDSe3Kz7BtqY63bT8z8UpzrtCu6VhdwE8b7GgB+EAjr6jfR4GAAAAAElFTkSuQmCC"" />"
$_ht_img["domain"]     = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABMAAAAUCAIAAADgN5EjAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAD3SURBVDhP1ZOhDoMwEIbbheAwBDEJAkNQOAQvgOVpeAIeY0/AUyBwBIVBgOQRCCFhV+7oWDcx2Mw+0fz/XS9trlc+jqOmaZxz9sw8z7BCCq1kWRZMicqyLIuiwIQkTVNYsyxDK4miKAxDEBf0J1Ar4zi+rtxWUEOQ0jvEbbuua9sW/TRNeZ6jliRJous6atd1HccBIRowDENd12uceZ6HQkFuMAzjUbknCALTNMls2LbdNA2ZDbWyqqq3tyW143e9/Zz/qhS9hceFV0Lf9z0KBTlGMFIoxJmWZfkbGH2F0r4PmzFy/rZibkke5Isz4Y+TPMjZMxm7A32GTQOptCUXAAAAAElFTkSuQmCC"" />"
$_ht_img["list"]       = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABMAAAATCAYAAAByUDbMAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAC1SURBVDhP7ZSxCsMgFEVvS7eAP+Ag2cwUyHe4+Z1uzpkyStaQKTjkEwJOpu1D2lCaoa1DoT2DHu7weHLBQwhhRSaO6c4CbWathTEmRa+htYZSijz/ZtM0YRgGCoQQ8N6Tb9nLq6pCWZbkp+sxzzOccxTEGNH3PfmWvZwxdhuW9Zk/Moza7LoObdtS0DQNlmUh37LXZl3XkFKSU5uPjOOY7E5RFE9zznmyf5vv8MX/2Xoh+YcAZ6uBS5QbRMFoAAAAAElFTkSuQmCC"" />"
$_ht_img["warning"]    = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAATCAYAAACZZ43PAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAOkSURBVDhPZZNtTNVlGMZ/58+LwNDDWyAvY0jJlnFQDFFwiwiTms6VfbEyHU1rrbG1tYo5v1Sb6Ypqa21JpFaLEEd9aG1UvAgUlkDBkESE8RpCcA5wDnA4h8O5uw+coJdre/Y8/+d/3dd93c9zPyZR8B9MTU1TVl6NYYKo6Ehs1gleefkEAYHBfsY6DP+8hoar13njTCWb47Moee0EIWEW3JLN8ZNn6ejo9LPWseZg/M4k739QBUHbyN89Rbi3ipj0K/R3voR4Zmjr/5Ceniq2JNo58+ZzGhHqC8NU9km1tLaPEbghkUfy00iNLSd95wXsw8nU93cRG/AwuXm12IZjuVxbwdjMvTQ2fk9y7ARFzzyA6amiUil+fgcpUVVsiholbOMiOG04HdN81tDJkewMzNGjELiMeGHsdha3xh+nZfQk3T+/hREQEsU98d+weetPhJmT1NQWHVGERno5mHkIc+okLG2ARQOTBxJj2tgRU0pMdCgejxvDvWTShdbjCoHlJRA3BHv55UdIsjTz+qkICPcdkw4VwGfQFY5Xllj2eDCizDHY7ZpheQjcg0oYAe8k9tkZZVsIkVS9K40yVMClv1Tf6Qpi0eHApHdoRJhNjPyZooFq1dkF8zfANoQlRQnsJS+7VUvYqAlUb15pquUy3c38gp3Q4ECMyLAZfh/atmINvS6cs7p2EbcJEuLc7Lpf0y4sgG77JqdTdUIP8MdIH2lbEzEK922nf0h3lws0k/8M1KfJLOzPuUZAtArPeHBp9gWlzTogIOIoHW115OZkYmRsv4+B/hv0zRbrGSg5UEeQ6k3DscPdWLX5nD4DGuiY0+o86YxazXrNvezJyVlt5Ufz7+KrZguexZiVKnzl/NAE+cfSOHs+F5dWMacObCpqJJ6jvuYyBQ9aCA83q2PFwvyc5B88J3eGroq1HrG3IL9WIoW7HpO+7w7Iza/1+3OkrvKQ9A6LxMVnS3t7qy9U1t7CO6UXtb2zOJpfg737VQxtC32M2DWzrwTbXAQ7n5ympORtAjy3uHip3Be26uBvFBQWS2O7S/qun5Kmj5DG80jNu0j1ewkyYXXJx5/WSUJSpgwO9vsjNLl/XsHAwKDs3lskHbdFJsZvSt2XT8hvTWXiWBIpv1Qn5sgMuVJV6Wev4l8CPvT09Ep2ztNSduFbmZwWqW3okuPPnpb4hCypqPjCz1rH/wR8sFmt8sKLpyXvoSOSlLxH9u0/LNdamv1//wmRvwBs6Sugk7kkjgAAAABJRU5ErkJggg=="" />"
$_ht_img["computer"]   = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAATCAIAAAAS8MqlAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAC2SURBVDhPY/z//z8D6YAJSpMIQLZlZGR8+vQJKoAX8PHxzZgxA8gAaYuKihIREREWFobI4QJv37598+bNsmXLgGwWiJCampqVlRWEjQscO3YMqA3CJtNvo9owAJnaoNEN5REBENHt4eHx4cMHsCAUPAEDGTCACoGBgIAAhIGSA4Cab9269e/fv5NgYA4GjIyM6urqcA0QgKKturr6/v37UA4SUFBQaGtrg3LAAEUb8YCeEcDAAABPs0wdC5WWAwAAAABJRU5ErkJggg=="" />"
$_ht_img["computer_d"] = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAAUCAIAAAAP9fodAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAE0SURBVDhPzVOhjoNAEIXjgOaaYEDeQUUNP4BAoiGU8BtILB+AQNXzB4gm1TUNgQRFwJNwqJq6IwTBDQxpUkGTnmjuidnHvHnMbmaXHIaBeB5v8/okxm7H47HrujnxECzL6roOZLRFUfQxAbUl/EywbRv4bPuagPISvieg7Y9n+1c2kiQZhlmtVhCBz9nHNoqiqqoyTZPnecuyLpfLer1GadEG/86yTFXVOI4VRTmfz4ZhtG2L6qINJut5HvI0TSHCqHzfx8w7LjgT5AhZlsuypGkarwVsEi5TURSojt222+3nPTiOg/1sNpu+7zVNgxqIwEVRnFxT6xugtGkaaJskyeFw2O/3WLPb7SDCacMwxMq7h3M6na7X6/xBEIIg5HkeBEFd15Ikua7rOA5Kr31vr7QRxC8RFJSmJVj1lgAAAABJRU5ErkJggg=="" />"
$_ht_img["user"]       = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAATCAYAAACdkl3yAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsIAAA7CARUoSoAAAAE6SURBVDhPrZKxioNAFEXvCiGQwsYiJF06+zR+glhIwN9ImV9Ia5sfkJQpA4IhbRqxsdDGQrSwkZCQRkh2d2YfgWEU4uIp9N03zBl4M1/fv2AAOkXP5xNVVeF+v0NVVUynUyiKQqsyrSK2ebvdIs9z6gC6rmOz2WAymVBHRDqCeV3XFSSMJEmw2+0oyUiioiiQpiklkTAMcbvdKIlIouv1SlU7dV1TJSKJ5vN551BHoxFmsxklEWmHpmkwTZOSiOM4GI/HlEQ6r3+/3+N8PuPxePDrtywLtm3TqszwD5L9oijC4XBAlmV8sY3FYoHVaoXlcinM8i3yPA/H45E3P8EwDKzX67eMf+M47iVhXC4XBEFAiUS+7/PQl9PpRBWJ2PP/D2VZUkWipml46Mvr9aJqwOv/G/kADCQCfgA1z4QfxDPMjgAAAABJRU5ErkJggg=="" />"
$_ht_img["user_d"]     = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAAUCAIAAAAP9fodAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAGGSURBVDhPnZKxisJAEIaTCxIsEsFWSATBIGItKGkEH0DwKYQk72Hnm9imt1AULBQEUZsgaGFAGyMS45/biSSe3Ol9xWb+nfyzw87yQRBwn/NF3w95cdrtdjsej57npdPpTCbD8zwlYjzbzufzYDA4nU5MZrPZWq0miiKTDxJNosRwOHx4gOu6k8mERIyEDQb8RyJit9uhYRIRCdvlcqEoyR82SZJ+XoAgCNgnEZGw4eqKxSKJiHK5DCeJiBcDWCwW6/UaDaOKpmmFQoEScWD7hfF4XK/X0bmu6xgM7QYBnYbVcZzlcolBs3JozPf9druNuFqtjkYjmG3bbjab2CHbdDrdbDbh7xG5XK7T6WAfcVj++6rgx2ARhFey3++fPECW5fl8nkqlWq0WJFbEs9mMZUPbarViIg5eWT6fv16vjUYDEitiRVFYNrQdDgcm4my3W9M0ERiGgaOwok/Lslg2tOHJMxEHtSuVSq/Xwwn9fl9V1W63WyqVWPbF3N4h8Ure5182jrsDnqnTiLVv0skAAAAASUVORK5CYII="" />"
$_ht_img["foreign"]    = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABMAAAAVCAIAAAAra0KGAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACBSURBVDhPY/z//z8DWYAJSpMOQHbu3r17/vz5UAFCIDEx0dXVFchgAeJPnz4BSQkJCRYWEBcX+PPnz4sXLyCKgQCh1NraGsrCDdauXQtlUeLPUZ2EwEjRiUh9d+7c4eDggHKwgR8/fkBZYADSCdFw8eJFsAgBADd9IHI2/XUyMAAAa/wlwrC6ezEAAAAASUVORK5CYII="" />"
$_ht_img["group"]      = "<img src=""data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABIAAAATCAIAAAAS8MqlAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAEqSURBVDhPvZG7ioNAFIb3NQXRwsrOTnyBdD6BNyzF0kBAsEmKRbw1gtpIioCFnaSy0D06Z5IxmIVssV8xHM78n+Oc+Zr/xFttmqa+79u2hRVq7FL2Ncidz2eDAvWLua/dbjc0KNDBvZV9rWkay7LQWCmKAvdWNhp88nA48DwPa5ZlruuiZBhwSQytPLVhGCRJ4ihQp2lKnCAI3t4tDEM0KKfTKYqiuq5/m+TxeMQ4BTq2bZMDfd+/Xq8YZbWu6wRBQIPjoIbrEedBVVUkvBlJkiSKooADK7wVnIBxiud5JInaOI5lWX5vieP4crmwL2GaJskv2v1+1zSN/NsLqqrmeQ6DefDUHMfB1B66rpMoy6LJsoyRPURRJFGWRWMHuAuJsuBIPuU/tXn+Aew0OJtgwbEnAAAAAElFTkSuQmCC"" />"
 
 
$_html = @"
<html>
<head>
    <title>Forest Report</title>
    <style type="text/css" title="Amaya theme">
        body {
            font-size: 12pt;
            font-family: Helvetica, Arial, sans-serif;
            font-weight: normal;
            font-style: normal;
            color: black;
            background-color: grey;
            line-height: 1.2em;
            margin-left: 2em;
            margin-right: 2em;
        }
        h1 {
            font-size: 170%;
            font-weight: bold;
            font-style: normal;
            font-variant: small-caps;
            text-align: left;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1em;
           }
        h2 {
            font-size: 130%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
        h3 {
            font-size: 110%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
    </style>
</head>
<body>
<h1>Forest Report</h1>
"@
 
# Collection information
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`" cellspacing=`"8`">`n"
$_html += "<tr><td colspan=`"2`">`n"
$_html += "<h2>$($_ht_img["forest"]) Forest: $($script:_ht_forest.Name)</h2>"
$_html += "</td></tr>`n"
$_html += "<tr><td>`n"
$_html += @"
    <h3>Collection Information</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Collection Date</b></td>
              <td>$_ScriptStartDate</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of domain</b></td>
              <td>$($_ht_sid_domain.Count)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of unique admins</b></td>
              <td>$($script:_ht_forest.TotalUniqueAdmins)</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td>`n"
$_html += "<td>`n"
$_html += @"
<h3>&nbsp;</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>User</b></td>
              <td>$_collection_user_domain\$_collection_user_name</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Computer</b></td>
              <td>$_collection_machine_name (Domain: $_collection_machine_domain)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Version</b></td>
              <td>$_ScriptVersion</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td></tr>`n"
 
 
#
$_html += "<tr><td>`n"
$_html += @"
    <h3>Technical Details</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Forest Installation</b></td>
              <td>$($script:_ht_forest.CreationTime)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Value</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristics)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Version</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristicsVersion)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Last Update</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristicsWhenModified)</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td>`n"
$_html += "<td>`n"
$_html += @"
<h3>&nbsp;</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Functional Mode</b></td>
              <td>$($script:_ht_forest.Mode)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>tombstoneLifetime</b></td>
              <td>$($script:_ht_forest.TSLz)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>tombstoneLifetime Version</b></td>
              <td>$($script:_ht_forest.TSLVersion)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>tombstoneLifetime Last Modification</b></td>
              <td>$($script:_ht_forest.TSLWhenModified)</td>
            </tr>
 
          </tbody>
        </table>
"@
$_html += "</td></tr>`n"
 
#
$_html += "</table>`n"
 
 
$_html += "<h1>Domain Summary</h1>"
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object {
    $_current_domain_sid = $_.Name
    $_html += "<table width=`"100%`" style=`"background-color:#ffffff`" cellspacing=`"8`">`n"
    $_html += "<tr><td colspan=`"2`">`n"
    $_html += "<h2>$($_ht_img["forest"]) Domain: $($_.Value.DomainName) [$_current_domain_sid]</h2></td></tr>`n"
    $_html += "<tr><td colspan=`"2`">`n"
    $_html += @"
    <h3>Additional Information</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>NetBIOS name</b></td>
              <td>$($_.Value.DomainNetbios)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Parent domain name</b></td>
              <td>$($_.Value.DomainParentName)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Domain Promotion Date</b></td>
              <td>$($_.Value.CreationTime)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Functional Level</b></td>
              <td>$($_FLPattern[$_.Value.DomainFL])</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Default Built-in Admin Account</b></td>
              <td>$($_.Value.DomainAdmin500) [UAC: $($_.Value.DomainAdmin500UAC)]</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Default Guest Account</b></td>
              <td>$($_.Value.DomainGuest501) [UAC: $($_.Value.DomainGuest501UAC)]</td>
            </tr>
          </tbody>
        </table>
 
"@
$_html += "<p>&nbsp;</p>"
$_html += "</td></tr>`n"
 
# limit for display set to 80
$_html_barlimitdisplay = 80
Write-Debug "Limit for the graphic: $_html_barlimitdisplay members"
    #Go for listing each group
    $_.Value.DomainGroups.GetEnumerator() | ForEach-Object {
        $_html += "<tr>`n"
        $_html += "<td align=`"right`" width=`"25%`">$($_.Value.ActualName)</td>`n"
        $_html += "<td align=`"left`" width=`"75%`">`n"
        $_current_stats = $script:_ht_sid_domain[ $_current_domain_sid ].DomainGroups[$_.Value.ActualName].Stats
        Write-Debug "Writing line for: $($_.Value.ActualName)"
        # Determine the color
        If ( $_.Value.LimitError -lt 0 ) {
            $_html_cellbackground = "#0080c0" #There is no limit -1 in the value
        } ElseIf ( $_current_stats["NestedMembersTotalNoGroups"] -gt $_.Value.LimitError ) {
            $_html_cellbackground = "#f00000"
        } ElseIf ( $_current_stats["NestedMembersTotalNoGroups"] -gt $_.Value.LimitWarning ) {
            $_html_cellbackground = "#ffd800"
        } Else {
            $_html_cellbackground = "#0080c0"
        }
        # Plurial management
        If ( $_current_stats["NestedMembersTotalNoGroups"] -gt 1 ) {
            $_html_cellplurial = "s"
        } Else {
            $_html_cellplurial = ""
        }
        $_html += "<table border=`"0`" cellpadding=`"0`" cellspacing=`"0`">`n"
        $_html += "<tr>`n"
 
        for ($_i = 0; $_i -lt $_current_stats["NestedMembersTotalNoGroups"] -and $_i -lt $_html_barlimitdisplay ; $_i++) { 
            $_html += "<td style=`"background-color:$_html_cellbackground`"><font color=`"$_html_cellbackground`">**</font></td>`n"
        }
        $_html += "<td style=`"background-color:#ffffff`">&nbsp;$($_current_stats["NestedMembersTotalNoGroups"])&nbsp;"
        $_html += "</td>`n"
        $_html += "</tr>`n"
        $_html += "</table>`n"
 
        $_html += "</td>`n"
        $_html += "</tr>`n"
    }
    $_html += "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>`n"
    $_html += "<tr>`n"
    $_html += "<td align=`"right`" width=`"25%`">Number of accounts:</td>`n"
 
    $_html += "<td>$($_.Value.TotalUniqueAdmins)</td>`n"
    $_html += "</tr>`n"
    $_html += "<tr><td>&nbsp;</td><td>&nbsp;</td></tr>`n"
    $_html += "<tr><td colspan=`"2`">"
    # Include raw data in TXT
    $_html += "</td></tr>`n"
    $_html += "</table>`n"
    $_html += "<p>&nbsp;</p>"
}
Write-Verbose "Generating HTML output for the schema analysis"
$_html += "<h1>Schema Information</h1>"
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`" cellspacing=`"8`">`n"
$_html += "<tr><td>`n"
$_html += "<h2>$($_ht_img["list"]) List of classes with a modified security descriptor</h2></td></tr>`n"
$_html += "<tr><td>`n"
    $_html += @"
    <h3>Additional Information</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Schema Level</b></td>
              <td>$($script:_ht_forest.SchemaLevel) - $($_SchemaPattern[ $script:_ht_forest.SchemaLevel ])</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Last Schema Update</b></td>
              <td>$($script:_ht_forest.SchemaLevelWhenModified)</td>
            </tr>
          </tbody>
        </table>
"@
 If ( $SkipSchema -eq $false ) 
{
    Write-Verbose "Listing all the classes with a modified defaultSecurityDescriptor for the HTML report"
    $script:_ht_forest.SchemaSecurity.GetEnumerator() | ForEach-Object `
    {
        #Write-Debug "Class: $($_.Name)"
        #Write-Debug "defaultSecurityDescriptorVersion: $($_.Value.defaultSecurityDescriptorVersion)"
        If ( $_.Value.defaultSecurityDescriptorVersion -gt 1 )
        {
            $_html += "<p>&nbsp;</p>"
            $_html += "<table border=`"0`" cellspacing=`"0`" cellpadding=`"2`">"
            $_html += "<col />"
            $_html += "<col />"
            $_html += "<tbody>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>Class LDAP name</b></td>"
            $_html += " <td width=`"75%`">$($_.Name)</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>defaultSecurityDescriptor version</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.defaultSecurityDescriptorVersion)</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>defaultSecurityDescriptor last modification</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.defaultSecurityDescriptorWhenModified)</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>defaultSecurityDescriptor</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.defaultSecurityDescriptor)</td>"
            $_html += "</tr>"
            $_html += "</tbody>"
            $_html += "</table>"
        }
    }
    $_html += "</td></tr>`n"
    $_html += "<tr><td>`n"
    $_html += "<h2>$($_ht_img["list"]) List of the attributes with a modified searchFlags</h2></td></tr>`n"
    $_html += "<tr><td>`n"
    Write-Verbose "Listing all the attributes with a modified searchFlags for the HTML report"
    $script:_ht_forest.SchemaAttributes.GetEnumerator() | ForEach-Object `
    {
        #Write-Debug "Attribute: $($_.Name)"
        #Write-Debug "searchFlags: $($_.Value.searchFlags)"
        If ( $_.Value.searchFlagsVersion -gt 1 )
        {
            $_html += "<p>&nbsp;</p>"
            $_html += "<table border=`"0`" cellspacing=`"0`" cellpadding=`"2`">"
            $_html += "<col />"
            $_html += "<col />"
            $_html += "<tbody>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>Attribute LDAP name</b></td>"
            $_html += " <td width=`"75%`">$($_.Name)</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>searchFlags</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.searchFlags)<br>"
            $_searchflags = $_.Value.searchFlags
            $_searchFlagsPattern.Keys | Where-Object { $_ -band $_searchflags } | ForEach-Object {
                $_html += "<code>+ $($_searchFlagsPattern.Get_Item($_)[0])</code><br />"
                $_html += "<code>  $($_searchFlagsPattern.Get_Item($_)[1])</code><br />"
            }
            $_html += "</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>searchFlags version</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.searchFlagsVersion)</td>"
            $_html += "</tr>"
            $_html += "<tr>"
            $_html += " <td style=`"background-color:#c6c6c6`" width=`"25%`"><b>searchFlags last modified</b></td>"
            $_html += " <td width=`"75%`">$($_.Value.searchFlagsWhenModified)</td>"
            $_html += "</tr>"
            $_html += "</tbody>"
            $_html += "</table>"
        }
    }
    $_html += "</td></tr>`n"
    $_html += "<tr><td>`n"
    $_html += "<h2>$($_ht_img["list"]) Listing all the sensitive attributes</h2></td></tr>`n"
    $_html += "<tr><td>`n"
    Write-Verbose "Listing all the sensitive attributes"
    $script:_ht_forest.SensitiveAttributes.GetEnumerator() | ForEach-Object `
                                    {
    $_html += "<strong>$($_.Name)</strong>`n"
    $_html += "<pre>"
    $_.Value | ForEach-Object `
    {
        $_html += $_
        $_html += "`n"
    }
    $_html += "</pre>`n"
    }
    $_html += "</td></tr>`n"
    $_html += "</table>`n"
}
$_html += "<p>&nbsp;</p>"
# Some stats on generation time for the sake of making stats...
$_ScriptHTLMEnd  = [System.DateTime]::Now
$_ScriptHTMLTime = $_ScriptHTLMEnd - $_ScriptHTLMStart
Write-Verbose "Report generation time: $_ScriptHTMLTime"
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`">`n"
$_html += "<tr><td><center>Collection time: $_ScriptCollectionTime</center></td></tr>`n"
$_html += "<tr><td><center>Analysis time: $_ScriptAccountAnalisysTime</center></td></tr>`n"
$_html += "<tr><td><center>Report generation time: $_ScriptHTMLTime</center></td></tr>`n"
$_html += "</table>`n"
$_html += "</body>`n"
$_html += "</html>`n"
$_html | Out-File "$($_timestamp)_Forest.html"
 
$_html = @"
<html>
<head>
    <title>Domains Report</title>
    <style type="text/css" title="Amaya theme">
        body {
            font-size: 12pt;
            font-family: Helvetica, Arial, sans-serif;
            font-weight: normal;
            font-style: normal;
            color: black;
            background-color: grey;
            line-height: 1.2em;
            margin-left: 2em;
            margin-right: 2em;
        }
        h1 {
            font-size: 170%;
            font-weight: bold;
            font-style: normal;
            font-variant: small-caps;
            text-align: left;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1em;
           }
        h2 {
            font-size: 130%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
        h3 {
            font-size: 110%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
       a[href] {
                color: white;
                text-decoration: none;
           }
    </style>
</head>
<body>
<h1>Domains Report</h1>
"@
 
# Collection information
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`" cellspacing=`"8`">`n"
$_html += "<tr><td colspan=`"2`">`n"
$_html += "<h2>$($_ht_img["forest"]) Forest: $($script:_ht_forest.Name)</h2>"
$_html += "</td></tr>`n"
$_html += "<tr><td>`n"
$_html += @"
    <h3>Collection Information</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Collection Date</b></td>
              <td>$_ScriptStartDate</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of domain</b></td>
              <td>$($_ht_sid_domain.Count)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of unique admins</b></td>
              <td>$($script:_ht_forest.TotalUniqueAdmins)</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td>`n"
$_html += "<td>`n"
$_html += @"
<h3>&nbsp;</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>User</b></td>
              <td>$_collection_user_domain\$_collection_user_name</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Computer</b></td>
              <td>$_collection_machine_name (Domain: $_collection_machine_domain)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Version</b></td>
              <td>$_ScriptVersion</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td></tr>`n"
$_html += "<tr><td>`n"
$_html += @"
    <h3>Technical Details</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Forest Installation</b></td>
              <td>$($script:_ht_forest.CreationTime)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Value</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristics)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Version</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristicsVersion)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>dsHeuristic Last Update</b></td>
              <td>$($script:_ht_forest.ForestdSHeuristicsWhenModified)</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td>`n"
$_html += "<td>`n"
$_html += @"
<h3>&nbsp;</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Account Operators protected</b></td>
              <td>$_admin_protected_ao</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Server Operators protected</b></td>
              <td>$_admin_protected_so</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Print Operators protected</b></td>
              <td>$_admin_protected_po</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Backup Operators protected</b></td>
              <td>$_admin_protected_bo</td>
            </tr>
 
          </tbody>
        </table>
"@
$_html += "</td></tr>`n"
$_html += "</table>`n<p></p>"
 
$script:_ht_sid_domain.GetEnumerator() | ForEach-Object {
    $_html += @"
    <table border="0" style="width: 100%" cellpadding="10" cellspacing="0">
      <col />
      <tbody>
        <tr>
          <td style="background-color:#ffffff"><h2>$($_ht_img["domain"]) Domain: $($_.Value.DomainName) [$_current_domain_sid]</h2>
 
 
            <h3>Group Membership</h3>
"@
    $_.Value.DomainGroups.GetEnumerator() | ForEach-Object {
        
        
    $_html += @"
            <h3>$($_.Name)</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Actual name</b></td>
              <td>$($_.Value.ActualName)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>SID</b></td>
              <td>$($_.Value.SID)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>AdminSDHolderProtected</b></td>
              <td>$($_.Value.AdminSDHolderProtected)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Admins</b></td>
              <td>$($_.Value.Admins)</td>
            </tr>
          </tbody>
        </table>
 
            $($_ht_img["group"]) <strong>$($_.Value.ActualDN)</strong><br>
"@
 
    $_.Value.ActualMembers.GetEnumerator() | ForEach-Object {
 
        $_i_level = ""
        $_i_img = "foreign"
        $_i_img_d = ""
        $_i_current_sid = $_.ActualMemberSID
 
        Switch ( $_.ActualMemberType ) {
            "C" { $_i_img = "computer" ; If ( $script:_ht_sid_objects[$_i_current_sid].userAccountControl -band 2 ) { $_i_img_d = "_d" } }
            "P" { $_i_img = "user" ; If ( $script:_ht_sid_objects[$_i_current_sid].userAccountControl -band 2 ) { $_i_img_d = "_d" }}
            "F" { $_i_img = "foreign" }
            "G" { $_i_img = "group" }
        }
        for ($_i_for = 0 ; $_i_for -le  $($_.ActualMemberLevel) ; $_i_for++) { $_i_level += "&mdash;"}
 
        $_i_primaryid = ""
        If ( $_.ActualMemberPrimary -eq $true ) { $_i_primaryid = " (via PrimaryGroupID)" }
        $_html += "<span style=""color:#ffffff"">$($_i_level)</span><a href=""#$_i_current_sid"" >$($_ht_img[$_i_img+$_i_img_d])</a> $($_.ActualMemberDn)$_i_primaryid<br />"
        #
    }
 
    $_html += "<PRE># Stats #`n"
    $_html += "Direct Membership`n"
    $_html += "      Group: $($_.Value.Stats.DirectMembersG)`n"                                                                                                                   
    $_html += "    Foreign: $($_.Value.Stats.DirectMembersF)`n"
    $_html += "       User: $($_.Value.Stats.DirectMembersP)`n"
    $_html += "   Computer: $($_.Value.Stats.DirectMembersC)`n"
    $_html += "      Total: $($_.Value.Stats.DirectMembersTotal)`n"
    $_html += "      Total: $($_.Value.Stats.DirectMembersTotalNoGroups) (without groups)`n"
    $_html += "Nested Membership`n"
    $_html += "      Group: $($_.Value.Stats.NestedMembersG)`n"                                                                                                                   
    $_html += "    Foreign: $($_.Value.Stats.NestedMembersF)`n"
    $_html += "       User: $($_.Value.Stats.NestedMembersP)`n"
    $_html += "   Computer: $($_.Value.Stats.NestedMembersC)`n"
    $_html += "      Total: $($_.Value.Stats.NestedMembersTotal)`n"
    $_html += "      Total: $($_.Value.Stats.NestedMembersTotalNoGroups) (without groups)</PRE>"
}
 
# Summary
$_html += "<p></p><h3>Groups Summary</h3><table border=""0"">"
 
<#
$_.Value.DomainSummary | ForEach-Object {
 
    $_html += "<tr><td style=""text-align:right"">$( $_.Split(";")[0] ): </td><td><strong>$( $_.Split(";")[1] )</strong></td></tr>"
 
}
 
#>
 
$_html += "</table>"
$_html += @"
            <p></p>
 
            <p></p>
          </td>
        </tr>
      </tbody>
    </table>
 
    <p></p>
"@
}
 
Write-Verbose "Generating Account Details HTMP output"
$_html += @"
<table border="0" style="width: 100%" cellpadding="10">
  <col />
  <tbody>
    <tr>
      <td style="background-color:#ffffff"><h2>$($_ht_img["list"]) Accounts</h2>
"@
 
$_i_objects = 0
$_i_objects_total = $_ht_sid_objects.Count
$_ht_sid_objects.GetEnumerator() | ForEach-Object {
 
    Write-Progress -Activity "Writing object's information.." -Status "$($_i_objects+1) / $_i_objects_total" -CurrentOperation "Writing information for $($_.Name).." -Id 0 -PercentComplete ( $_i_objects++ / $_i_objects_total * 100 )
    Write-Debug "Generating HTML output for: $($_.Value.distinguishedName) "
    $_i_img_d = ""
    $_i_account_uac = $_.Value.userAccountControl
    Switch ( $_.Value.objectCategory ) {
        "CN=Computer,$_forest_schema" { $_i_img = "computer" ; If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" } }
        "CN=Person,$_forest_schema" { $_i_img = "user" ; If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }}
        "CN=Foreign-Security-Principal,$_forest_schema"{ $_i_img = "foreign" }
    }
    $_html += "<a name=""$($_.Name)"">$($_ht_img[$_i_img+$_i_img_d])</a><b> $($_.Value.distinguishedName)</b><br />"
    $_UAC = $_.Value.userAccountControl
    $_UACPattern.Keys | Where-Object { $_ -band $_UAC } | ForEach-Object {
        $_html += "<code>+ $($_UACPattern.Get_Item($_)[0])</code><br />"
        $_html += "<code>  $($_UACPattern.Get_Item($_)[1])</code><br />"
    }
    $_html += "<table border=""0"">"
    $_html += "<tr><td style=""text-align:right"">objectCategory</td><td><strong>$($_.Value.objectCategory)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">objectSid</td><td><strong>$($_.Name)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">sAMAccountName</td><td><strong>$($_.Value.sAMAccountName)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">displayname</td><td><strong>$($_.Value.displayname)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userPrincipalName</td><td><strong>$($_.Value.userPrincipalName)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userPrincipalNameVersion</td><td><strong>$($_.Value.userPrincipalNameVersion)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userPrincipalNameTime</td><td><strong>$($_.Value.userPrincipalNameTime)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">accountExpires</td><td><strong>$($_.Value.accountExpires)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userAccountControl</td><td><strong>$($_.Value.userAccountControl)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">whenCreated</td><td><strong>$($_.Value.whenCreated)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">pwdLastSet</td><td><strong>$($_.Value.pwdLastSet)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">passwordVersion</td><td><strong>$($_.Value.passwordVersion)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">lastLogonTimestamp</td><td><strong>$($_.Value.lastLogonTimestamp)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">description</td><td><strong>$($_.Value.description)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">sidHistory</td><td><strong>$($_.Value.sidHistory)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userWorkstations</td><td><strong>$($_.Value.userWorkstations)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">lockoutTime</td><td><strong>$($_.Value.lockoutTime)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">mail</td><td><strong>$($_.Value.mail)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">userPassword</td><td><strong>$($_.Value.userPassword)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">adminCount</td><td><strong>$($_.Value.adminCount) ($($_.Value.adminCountDate))</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">admin member</td><td><strong>$($_.Value.admin)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">AdminSDHolderProtected</td><td><strong>$($_.Value.AdminSDHolderProtected)</strong></td></tr>"
    $_html += "<tr><td style=""text-align:right"">Collected Groups</td><td>"
    $_.Value.membership | Group-Object | ForEach-Object `
    {
        <#
        What's the... this membership thingy needs to be corrected
        If ( $_.Count -gt 1 )
        {
            $_groupcount = "$($_.Count) times"
        }
        #>
        $_html += "$($_.Name)<br>`n" 
    }
    $_html += "</td></tr>"
    $_html += "</table><p></p>"
}
 
$_html += @"
      </td>
    </tr>
  </tbody>
</table>
"@
 
$_html += "</table>`n"
$_html += "<p>&nbsp;</p>"
# Some stats on generation time for the sake of making stats...
$_ScriptHTLMEnd  = [System.DateTime]::Now
$_ScriptHTMLTime = $_ScriptHTLMEnd - $_ScriptHTLMStart
Write-Verbose "Report generation time: $_ScriptHTMLTime"
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`">`n"
$_html += "<tr><td><center>Collection time: $_ScriptCollectionTime</center></td></tr>`n"
$_html += "<tr><td><center>Analysis time: $_ScriptAccountAnalisysTime</center></td></tr>`n"
$_html += "<tr><td><center>Report generation time: $_ScriptHTMLTime</center></td></tr>`n"
$_html += "</table>`n"
$_html += "</body>`n"
$_html += "</html>`n"
$_html | Out-File "$($_timestamp)_Domains.html"
 
Write-Verbose "Generating Accounts the HTML report"
$_html = @"
<html>
<head>
    <title>Accounts Report</title>
    <style type="text/css" title="Amaya theme">
        body {
            font-size: 12pt;
            font-family: Helvetica, Arial, sans-serif;
            font-weight: normal;
            font-style: normal;
            color: black;
            background-color: grey;
            line-height: 1.2em;
            margin-left: 2em;
            margin-right: 2em;
        }
        h1 {
            font-size: 170%;
            font-weight: bold;
            font-style: normal;
            font-variant: small-caps;
            text-align: left;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1em;
           }
        h2 {
            font-size: 130%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
        h3 {
            font-size: 110%;
            font-weight: bold;
            font-style: normal;
            padding: 0;
            margin-top: 1em;
            margin-bottom: 1.1em;
           }
       a[href] {
                color: white;
                text-decoration: none;
           }
    </style>
</head>
<body>
<h1>Accounts Report</h1>
"@
 
# Collection information
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`" cellspacing=`"8`">`n"
$_html += "<tr><td colspan=`"2`">`n"
$_html += "<h2>$($_ht_img["forest"]) Forest: $($script:_ht_forest.Name)</h2>"
$_html += "</td></tr>`n"
$_html += "<tr><td>`n"
$_html += @"
    <h3>Collection Information</h3>
    <table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>Collection Date</b></td>
              <td>$_ScriptStartDate</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of domain</b></td>
              <td>$($_ht_sid_domain.Count)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Number of unique admins</b></td>
              <td>$($script:_ht_forest.TotalUniqueAdmins)</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td>`n"
$_html += "<td>`n"
$_html += @"
<h3>&nbsp;</h3>
<table border="0" cellspacing="0" cellpadding="2">
          <col />
          <col />
          <tbody>
            <tr>
              <td style="background-color:#c6c6c6"><b>User</b></td>
              <td>$_collection_user_domain\$_collection_user_name</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Computer</b></td>
              <td>$_collection_machine_name (Domain: $_collection_machine_domain)</td>
            </tr>
            <tr>
              <td style="background-color:#c6c6c6"><b>Version</b></td>
              <td>$_ScriptVersion</td>
            </tr>
          </tbody>
        </table>
"@
$_html += "</td></tr>`n"
$_html += "</table>`n<p></p>"
 
$_html += @"
<p></p>
<table border="0" style="width: 100%" cellpadding="10">
  <col />
  <tbody>
    <tr>
      <td style="background-color:#ffffff"><h2>$($_ht_img["list"]) Admin Accounts with Email</h2>
"@
 
 
$_html += "<table border=""0"" cellspacing=""0"">"
$_ht_check_account["MAIL"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.mail)</td><td>$($_.distinguishedName)</td></tr>"
}
$_html += "</table>"
$_html += "<br /><b>Admins Accounts with email: $($_ht_check_account["MAIL"].Count)</b>"
$_html += "<p></p>"
$_html += "<h2>$($_ht_img["list"]) Admin Accounts with modified UPN</h2>"
 
$_html += "<table border=""0"" cellspacing=""0"">"
    $_html += "<tr><td></td>"
    $_html += "<td><strong>UPN</strong></td>"
    $_html += "<td><strong>Modified</strong></td>"
    $_html += "<td><strong>Last Modified Time</strong></td>"
    $_html += "</tr>"
$_ht_check_account["UPN"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.userPrincipalName)</td>"
    $_html += "<td>$($_.userPrincipalNameVersion)</td>"
    $_html += "<td>$($_.userPrincipalNameTime)</td>"
    $_html += "</tr>"
}
$_html += "</table>"
$_html += "<br /><b>Admins Accounts with modified UPN: $($_ht_check_account["UPN"].Count)</b>"
$_html += "<p></p>"
 
<#
$_account_stale_threshold    = 180
#>
$_html += "<h2>$($_ht_img["list"]) Stale Admin Accounts based on lastLogonTimestamp</h2>"
$_html += "<br><i>Threshold: $_account_stale_threshold days</i>"
 
$_html += "<table border=""0"" cellspacing=""0"">"
    $_html += "<tr><td></td>"
    $_html += "<td><strong>lastLogonTimestamp</strong></td>"
    $_html += "<td><strong>Delta</strong></td>"
    $_html += "<td><strong>distinguishedName</strong></td>"
    $_html += "</tr>"
$_ht_check_account["STALE"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.lastLogonTimestamp)</td>"
    $_html += "<td>$($_.Delta)</td>"
    $_html += "<td>$($_.distinguishedName)</td>"
    $_html += "</tr>"
}
$_html += "</table>"
$_html += "<br /><b>Stale Admins Accounts: $($_ht_check_account["STALE"].Count)</b>"
 
 
<#
$_account_stale_threshold    = 180
#>
$_html += "<h2>$($_ht_img["list"]) Stale Admin Accounts based on pwdLastSet</h2>"
$_html += "<br><i>Threshold: $_account_password_threshold days</i>"
 
$_html += "<table border=""0"" cellspacing=""0"">"
    $_html += "<tr><td></td>"
    $_html += "<td><strong>pwdLastSet</strong></td>"
    $_html += "<td><strong>Delta</strong></td>"
    $_html += "<td><strong>distinguishedName</strong></td>"
    $_html += "</tr>"
$_ht_check_account["PASSWORD_AGE"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.pwdLastSet)</td>"
    $_html += "<td>$($_.Delta)</td>"
    $_html += "<td>$($_.distinguishedName)</td>"
    $_html += "</tr>"
}
$_html += "</table>"
$_html += "<br /><b>Stale Password Admins Accounts: $($_ht_check_account["PASSWORD_AGE"].Count)</b>"
 
 
$_html += "<h2>$($_ht_img["list"]) Admins Accounts with a password that never expires</h2>"
 
$_html += "<table border=""0"" cellspacing=""0"">"
    $_html += "<tr><td></td>"
    $_html += "<td><strong>pwdLastSet</strong></td>"
    $_html += "<td><strong>distinguishedName</strong></td>"
    $_html += "</tr>"
$_ht_check_account["PASSWORD_NOEXPIRE"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.pwdLastSet)</td>"
    $_html += "<td>$($_.distinguishedName)</td>"
    $_html += "</tr>"
}
$_html += "</table>"
$_html += "<br /><b>Admins Accounts with a password that never expires: $($_ht_check_account["PASSWORD_NOEXPIRE"].Count)</b>"
 
 
$_html += "<h2>$($_ht_img["list"]) Admins Accounts with a password that never changed</h2>"
 
$_html += "<table border=""0"" cellspacing=""0"">"
    $_html += "<tr><td></td>"
    $_html += "<td><strong>whenCreated</strong></td>"
    $_html += "<td><strong>pwdLastSet</strong></td>"
    $_html += "<td><strong>distinguishedName</strong></td>"
    $_html += "</tr>"
$_ht_check_account["PASSWORD_NEVER"].GetEnumerator() | ForEach-Object `
{
    $_i_account_uac = $_.userAccountControl
    $_i_img_d = ""
    If ( $_i_account_uac -band 2 ) { $_i_img_d = "_d" }
    $_html += "<tr><td>$($_ht_img["user"+$_i_img_d])</td>"
    $_html += "<td>$($_.whenCreated)</td>"
    $_html += "<td>$($_.pwdLastSet)</td>"
    $_html += "<td>$($_.distinguishedName)</td>"
    $_html += "</tr>"
}
$_html += "</table>"
$_html += "<br /><b>Admins Accounts with a password that never changed: $($_ht_check_account["PASSWORD_NEVER"].Count)</b>"
 
$_html += @"
        <p></p>
      </td>
    </tr>
  </tbody>
</table>
 
<p></p>
"@
 
Write-Verbose "Report generation time: $_ScriptHTMLTime"
$_html += "<table width=`"100%`" style=`"background-color:#ffffff`">`n"
$_html += "<tr><td><center>Collection time: $_ScriptCollectionTime</center></td></tr>`n"
$_html += "<tr><td><center>Analysis time: $_ScriptAccountAnalisysTime</center></td></tr>`n"
$_html += "<tr><td><center>Report generation time: $_ScriptHTMLTime</center></td></tr>`n"
$_html += "</table>`n"
$_html += "</body>`n"
$_html += "</html>`n"
$_html | Out-File "$($_timestamp)_Accounts.html"
 
 
$_ScriptEndDate = [System.DateTime]::Now
Write-Verbose "Script finish time: $_ScriptEndDate"
$_ScriptExecTime = $_ScriptEndDate - $_ScriptStartDate
Write-Debug "Total execution time: $_ScriptExecTime"
 
# Stop the transcript
Stop-Transcript 

 