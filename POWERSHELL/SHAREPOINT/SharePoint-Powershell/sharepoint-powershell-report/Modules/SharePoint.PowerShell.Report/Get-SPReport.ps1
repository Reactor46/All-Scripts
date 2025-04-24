function Get-SPReport
{
  param
  (
    [parameter(Mandatory = $true)]
    [string]$ExportPath,
    [switch]$FromExportedFiles,
    [string]$ReportFilePath
  )
	
  if (-not $FromExportedFiles)
  {
    Export-SPRObjects -Path $ExportPath
  }
  $SPRObjects = Get-SPRObject -Path $ExportPath
	

  # 2	FARM OVERVIEW
  $title = 'Farm Overview'
  Write-Host -Object "Building section: $title"
  $report = Get-SPRFarmOverview -spfarm $SPRObjects.SPFarm -SPDatabases $SPRObjects.SPDatabase | 
  ConvertTo-EnhancedHTMLFragment -As List `
  -PreContent "<h2>$title</h2>" `
  -Properties FarmName, BuildVersion, SharePointLicense, ConfigurationDatabase, FarmServerCount

  # 2.1	SERVERS IN FARM                                    
  $title = 'Servers in Farm'  
  Write-Host -Object "Building section: $title"                               
  $report += Get-SPRServersInFarm -spserver $SPRObjects.SPServer | 
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServerName, Role, OperatingSystem, Memory 
  
  # 2.2	WEB APPLICATIONS AND SITE COLLECTIONS                                 
  $title = 'Web Applications and Site Collections' 
  Write-Host -Object "Building section: $title"                                
  $report += Get-SPRWebApplicationsAndSiteCollections -spsite $SPRObjects.SPSite |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, SiteCollection, SiteAdmins
      
  # 2.3	CONTENT DATABASES                                 
  $title = 'Content Databases'     
  Write-Host -Object "Building section: $title"                            
  $report += Get-SPRContentDatabases -SPDatabase $SPRObjects.SPDatabase |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Name
    
  # 3	FARM TOPOLOGY
  $title = 'Farm Topology' 
  Write-Host -Object "Building section: $title"                                
  $report += Get-SPRFarmTopology -spserver $SPRObjects.SPServer -spserviceinstance $SPRObjects.SPServiceInstance |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Name, IPAddress, Role, ContralAdministration   
       
  # 4	SITE TOPOLOGY                         
  $title = 'Site Topology'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSiteTopology -spsite $SPRObjects.SPSite |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, SiteCollection, ContentDatabase     
          
  #	5.1	SEARCH SERVICE APPLICATION                        
  $title = 'Search Topology'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchServiceApplicationTopology -SPEnterpriseSearchComponent $SPRObjects.SearchComponent|
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties AdminComponent, AnalyticsProcessingComponent, ContentProcessingComponent, CrawlComponent, IndexComponent, QueryProcessingComponent
      
  # 6	SERVICES ON SERVER                        
  $title = 'Services on Server'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRServicesOnServer -SPServer $SPRObjects.SPServer -SPServiceInstance $SPRObjects.SPServiceInstance |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties *

  # 7.1.1	WEB APPLICATIONS LIST
  $title = 'Web Application List'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationList -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, Url

  # 7.1.2	GENERAL SETTINGS
  $title = 'Web Application General Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationGeneralSettings -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, DefaultTimeZone, DefaultQuotaTemplate

  # 7.1.3	RESOURCE THROTTLING   
  $title = 'Web Application Resource Throttling'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationResourceThrottling -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, MaxItemsPerThrottledOperation, AllowOMCodeOverrideThrottleSettings, MaxQueryLookupFields, UnthrottledPrivilegedOperationWindowEnabled, MaxUniquePermScopesPerList, IsBackwardsCompatible, Days

  # 7.1.4	WORKFLOW SETTINGS
  $title = 'Web Application Workflow Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationWorkflowSettings -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, AllowInternalUsers, AllowExternalUsers
      
  # 7.1.5	WEB APPLICATION FEATURES                
  $title = 'Web Application Features'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationFeatures -SPFeature $SPRObjects.SPFeature -SPwebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, FeatureName, CompatibilityLevel, IsActive
      
  # 7.1.6	MANAGED PATHS
  $title = 'Web Application Managed Paths'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationManagedPaths -SPManagedPath $SPRObjects.SPManagedPath |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplicationUrl, Name, PrefixType

  # 7.1.7	SELF-SERVICE SITE CREATION
  $title = 'Web Application Self Service Site Creation'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationSelfServiceSiteCreation -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplicationName, WebApplicationUrl, Allowed, RequireSecondaryContact 
      
  # 7.1.8	WEB PART SECURITY
  $title = 'Web Application Web Part Security'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationWebPartSecurity -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, OnlineWebPartGallery, WebPartConnections, ScriptableWebParts 
      
  # 7.1.9	USER POLICY
  $title = 'Web Application User Policy'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationUserPolicy -SPWebApplication $SPRObjects.SPWebApplication -SPManagedAccount $SPRObjects.SPManagedAccount |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, UserDisplayName, Username, PolicyRoleBindings, SystemUser

  # 7.1.10	ALTERNATE ACCESS MAPPINGS
  $title = 'Web Application Alternate Access Mappings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationAlternateAccessMappings -SPAlternateUrl $SPRObjects.SPAlternateURL |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties *
  
  # 7.1.11	IIS SETTINGS
  $title = 'IIS Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRIISSettings -SPWebApplication $SPRObjects.SPWebApplication -SPAlternateURL $SPRObjects.SPAlternateURL -SPAuthenticationProvider $SPRObjects.SPAuthenticationProvider |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, Url, Zone, Authentication, ApplicationPoolName, ApplicationPoolIdentity, SSL, ClaimsAuth, CEIP
  
  # 7.1.12	OBJECT CACHE ACCOUNTS
  $title = 'Web Application Object Cache Accounts'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRWebApplicationObjectCacheAccounts -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, SuperUserAcount, SuperUserHasFullControl, SuperReaderAcount, SuperReaderHasFullRead
  
  # 7.2.1	SITE COLLECTIONS LIST
  $title = 'Site Collections List'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSiteCollectionList -SPSite $SPRObjects.SPSite |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, SiteCollectionTitle, SiteCollectionUrl, SiteCollectionContentDatabase, Owners
  
  # 7.2.2	USAGE AND PROPERTIES
  $title = 'Site Collections Usage and Properties'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSiteUsageAndProperties -SPSite $SPRObjects.SPSite |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties SiteCollectionTitle, SiteCollectionUrl, Language, Template, NumberOfSites, UIVersion, Storage

  # 7.2.3	QUOTA TEMPLATES
  $title = 'Quota Templates'
  Write-Host -Object "Building section: $title"
  $report += Get-SPRQuotaTemplates -SPQuotaTemplate $SPRObjects.SPQuotaTemplate |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties TemplateName, StorageMaximumLevel, StorageWarningLevel, InvitedUserMaximumLevel, UserCodeMaximumLevel, UserCodeWarningLevel
  
  # 7.2.4	SITE COLLECTION QUOTAS
  $title = 'Site Collection Quotas'
  Write-Host -Object "Building section: $title"
  $report += Get-SPRSiteCollectionQuotas -SPSite $SPRObjects.SPSite -SPQuotaTemplate $SPRObjects.SPQuotaTemplate |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties SiteCollection, URL, QuotaName, LockStatus, StorageMaximumLevel, StorageWarningLevel, UsageStorage
  
  # 7.3.1.1	DISTRIBUTED CACHE CONFIG
  $title = 'DISTRIBUTED CACHE CONFIG'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRDistributedCacheConfig -SPCacheClusterConfig $SPRObjects.SPCacheClusterConfig -SPCacheHost $SPRObjects.SPCacheHost |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Host, Size, CachePort, ClusterPort, ReplicationPort, ArbitrationPort, Status

  # 7.3.2.1	SEARCH SETTINGS
  $title = 'Search Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchSettings -SPEnterpriseSearchServiceApplication $SPRObjects.SPEnterpriseSearchServiceApplication -SPEnterpriseSearchServiceInstance $SPRObjects.SPEnterpriseSearchServiceInstance |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, DefaultContentAccessAccount, ContactEmailAddress, IndexLocation, SearchAlertsStatus, QueryLogging
  
  # 7.3.2.2	CONTENT SOURCES
  $title = 'Content Sources'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchContentSources -SPEnterpriseSearchServiceApplication $SPRObjects.SPREnterpriseSearchTopologyConfiguration |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, ContentSourceName, ContentSourceType
  
  # 7.3.2.3	START ADDRESSES
  $title = 'Start Addresses'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchContentSourcesStartAddresses -SPEnterpriseSearchServiceApplication $SPRObjects.SPREnterpriseSearchTopologyConfiguration |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, ContentSourceName, ContentSourceStartAddress
  
  # 7.3.2.4	CRAWL STORES
  $title = 'Crawl Stores'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchCrawlStores -SPEnterpriseSearchServiceApplication $SPRObjects.SPEnterpriseSearchServiceApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, Name, SQLServer
  
  # 7.3.2.5	FILE TYPES
  $title = 'File Types'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchFileTypes -SPEnterpriseSearchServiceApplication $SPRObjects.SPEnterpriseSearchServiceApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, FileTypes

  # 7.3.2.6	AUTHORITATIVE PAGES
  $title = 'Authoritative Pages'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchAuthoritativePages -SPEnterpriseSearchServiceApplication $SPRObjects.SPEnterpriseSearchServiceApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, Url, Level, Status
  
  # 7.3.2.7	FEDERATED LOCATIONS
  $title = 'Federated Locations'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchFederatedLocations -SPEnterpriseSearchServiceApplication $SPRObjects.SPEnterpriseSearchServiceApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, DisplayName, Author, Version, LocationType, Trigger, CreationDate
  
  # 7.3.2.8	SCOPES
  
  # 7.3.2.9	MANAGED PROPERTIES
  
  # 7.3.2.10	SEARCH TOPOLOGIES
  $title = 'Search Topologies'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRSearchTopologies -SPREnterpriseSearchTopologyConfiguration $SPRObjects.SPREnterpriseSearchTopologyConfiguration |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, Topology, CreationDate, State, ComponentCount, IsActive
  
  # 7.3.2.11	SEARCH COMPONENTS
  
  # 7.3.2.12	INDEX PARTITIONS
  
  # 7.3.3.1	MYSITE SETTINGS  
  $title = 'Mysite Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRMySiteSettings -SPServiceApplicationUserProfiles $SPRObjects.SPServiceApplicationUserProfiles -SPUserProfileManager $SPRObjects.SPUserProfileManager |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, Multilingual, ReadPermissionLevel, SenderEmailAddress, EnableNetbiosdomainName, MySiteManagedPath

  # 7.3.3.2	AUDIENCES
  $title = 'Audiences'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRMySiteSettings -SPServiceApplicationUserProfiles $SPRObjects.SPServiceApplicationUserProfiles -SPUserProfileManager $SPRObjects.SPUserProfileManager |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties *
  
  # 7.3.3.3	USER PROFILE PROPERTIES
  $title = 'User Profile Properties'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRUserProfuleProperties -SPUserProfileManager $SPRObjects.SPUserProfileManager |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DisplayName, PropertyType, DefaultPolicy, PrivacyPolicy, UserOverridePolicy, Replicable, Multivalue, Alias
  
  # 7.3.3.4	USER PROFILES COUNT
  $title = 'User Profile Count'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRUserProfuleCount -SPServiceApplicationUserProfiles $SPRObjects.SPServiceApplicationUserProfiles -SPUserProfileManager $SPRObjects.SPUserProfileManager |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties ServiceApplication, UserProfileCount
  
  # 7.3.4	SERVICE APPLICATIONS LIST
  
  # 7.3.5	APPLICATION PROXIES
  
  # 7.3.6	ADMINISTRATORS
  
  # 7.3.7	PERMISSIONS
  
  # 7.3.8	ASSOCIATIONS
  
  # 7.3.9	PROXY GROUPS
  
  # 7.4.1	CONTENT DATABASES
  $title = 'Content Databases'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRContentDatabases -SPDatabase $SPRObjects.SPDatabase |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Name, Status, Server, WebApplication, CurrentSiteCount, WarningSiteCount, MaximumSiteCount, AuthenticationType, ReadOnly, RecoveryModel, TotalSize 
  
  # 7.4.2	DATA RETRIEVAL SERVICES
  $title = 'DATA RETRIEVAL SERVICES'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRDataRetrievalServices -SPWebApplications $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, InheritTheGlobalSettings, EnableTheseDataRetrievalServices, EnableUpdateQuerySupport, EnableTheseDataSourceControles, ResponseSizeLimit, RequestTimeOut 
  
  # 7.4.3	OTHER DATABASES
  
  # 7.4.4	DATABASE FILES
  
  # 8.1.1	FARM FEATURES
  $title = 'Farm Features'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRFarmFeatures -SPWebServicesAdministration $SPRObjects.SPWebServicesAdministration -SPFarmFeatures $SPRObjects.SPFeature |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Title, ID, SolutionID, SolutionName, Version, Active, Hidden, Custom
  
  # 8.1.2	USER SOLUTIONS
  
  # 8.1.3	CROSS FIREWALL ACCESS ZONE
  $title = 'Web Application Cross FireWall Access Zone'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRCrossFirewallAccessZone -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, CrossFirewallAccessZone
  
  # 8.2.1	INCOMING EMAIL SETTINGS
  
  # 9.1.1	PROBLEMS
  
  # 9.2.1	JOB DEFINITIONS
  
  # 9.3.1.1	DIAGNOSTIC LOGGING SETTINGS
  $title = 'Diagnostic Logging Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRDiagnosticLoggingSettings -SPDiagConfig $SPRObjects.SPDiagnosticLogging |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties AllowLegacyTraceProviders, AppAnalyticsAutomaticUploadEnabled, CustomerExperienceImprovementProgramEnabled, DaysToKeepLogs, DownloadErrorReportingUpdatesEnabled, ErrorReportingAutomaticUploadEnabled, ErrorReportingEnabled, EventLogFloodProtectionEnabled, EventLogFloodProtectionNotifyInterval, EventLogFloodProtectionQuietPeriod, EventLogFloodProtectionThreshold, EventLogFloodProtectionTriggerPeriod, LogCutInterval, LogDiskSpaceUsageGB, LogLocation, LogMaxDiskSpaceUsageEnabled, ScriptErrorReportingDelay, ScriptErrorReportingEnabled, ScriptErrorReportingRequireAuth


  
  # 9.3.1.2	EVENT THROTTLING
  $title = 'Event Throttling'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPREventThrottling -SPLogLevels $SPRObjects.SPLogLevel |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Name, EventSeverity, TraceSeverity
  
  # 9.3.2	USAGE AND HEALTH DATA COLLECTION
  
  # 10.1	BACKUP SETTINGS
  
  # 11.1	GENERAL SETTINGS - USER LICENSING
  
  # 12.1.1	FARM ADMINISTRATORS
  $title = 'Farm Administrators'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRFarmAdministrators -SPFarmAdmins $SPRObjects.SPFarmAdmins |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Login, DisplayName
  
  # 12.2.1	MANAGED ACCOUNTS
  
  # 12.2.1.1	MANAGED ACCOUNT COMPONENTS
  
  # 12.2.2	SERVICE ACCOUNTS
  
  # 12.2.3	PASSWORD CHANGE SETTINGS
  
  # 12.2.4	AUTHENTICATION PROVIDERS
  
  # 12.2.5	ANTIVIRUS
  $title = 'Antivirus Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRFarmAntivirusSettings -SPWebServices $SPRObjects.SPWebServices |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties DownloadScanEnabled, UploadScanEnabled, AllowDownload, CleaningEnabled, Timeout, NumberOfThreads
  
  # 12.2.6	BLOCKED FILE TYPES
  
  # 13.1	INSTALLED PRODUCTS
  
  # 13.2	PATCH STATUS
  
  # 14.1	EXTERNAL SERVICE CONNECTIONS
  
  # 14.1.1	DOCUMENT CONVERSIONS
  $title = 'Document Conversions'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRDocumentConversions -SPWebApplication $SPRObjects.SPWebApplication |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties WebApplication, Converter, Type, Timeout, MaxRetries, MaxFileSize
  
  # 14.2.1	FORM TEMPLATES
  
  # 14.2.2	CONFIGURATION
  
  # 14.2.3	PROXY CONFIGURATION
  
  # 14.3.1	CONTENT DEPLOYMENT SETTINGS
  $title = 'Content Deployment Settings'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRContentDeploymentSettings -SPContentDeploymentSettings $SPRObjects.SPContentDeploymentSettings|
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties AcceptIncomingJobs, ExportWebServer, ImportWebServer, TemporaryFolder, RequireSecureConnection
  
  # 14.4	SHAREPOINT DESIGNER SETTINGS
  
  # 15.1	APP URLS
  
  # 16.1	LOCAL ADMINS
  
  # 16.2	PROCESSORS INFO
  
  # 16.3	DISKS LIST
  
  # 16.4	CERTIFICATES
  
  # 17.1	GENERAL SETTINGS - SQL
  
  # 17.2	MEMORY - SQL
  
  # 17.3	PROCESSORS INFO - SQL
  
  # 17.3.1	PROCESSORS AFFINITY - SQL
  
  # 17.4	SECURITY
  
  # 17.5.1	DEFAULT CONNECTION OPTIONS
  
  # 17.6	DATABASE SETTINGS
  
  # 17.7	ADVANCED SETTINGS
  
  # 17.8	SQL INFO
  
  # 17.8	SQL ALIASES
  
  # 17.10	SQL SERVICE ACCOUNTS
  
  # 18.1	ALL ACCOUNTS
  
  # 18.2	PRODUCT KEYS
  
  # 18.3	PASSPHRASE
  $title = 'Farm Passphrase'    
  Write-Host -Object "Building section: $title"                             
  $report += Get-SPRFarmPassphrase |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties Passphrase
  
  # 19	SCHEDULED TASKS
  $title = 'Scheduled Tasks'    
  Write-Host -Object "Building section: $title"
  $report += Get-SPRScheduledTasks -ScheduledTasks $SPRObjects.SPTasks |
  ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
  -PreContent "<h2>$title</h2>" `
  -Properties TaskName, State, Description, Source, URI, Version
  
  
 


  

  <#

      $title = 'SearchTopologies'
      Write-Host -Object "Building section: $title"
      $report += Get-SPRSearchTopologies -Topologies $SPRObjects.ActiveSearchTopology |
      ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
      -PreContent "<h2>$title</h2>" `
      -Properties ServiceApplication, Topology, CreationDate, State, ComponentCount, IsActive

      $title = 'SearchTopologies'
      Write-Host -Object "Building section: $title"
      $report += Get-SPRSearchTopologies -Topologies $SPRObjects.ActiveSearchTopology |
      ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
      -PreContent "<h2>$title</h2>" `
      -Properties ServiceApplication,Topology,CreationDate,State,ComponentCount,IsActive

      $title = 'SearchComponent'
      Write-Host -Object "Building section: $title"
      $report += Get-SPRSearchComponents -SearchComponents $SPRObjects.SearchComponent |
      ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
      -PreContent "<h2>$title</h2>" `
      -Properties Name,Server,ServiceApplication,Topology,IsActive

      $title = 'SearchIndex'
      Write-Host -Object "Building section: $title"
      $report += Get-SPRSearchIndexPartitions -IndexPartitions $SPRObjects.SearchServiceApplication |
      ConvertTo-EnhancedHTMLFragment -As Table -EvenRowCssClass 'even' -OddRowCssClass 'odd' -TableCssClass 'report-table' `
      -PreContent "<h2>$title</h2>" `
      -Properties Name,Server,ServiceApplication,Topology,RootDirectory,IndexLocation,Role,Service,Status
  #>
  #endregion
                 
  $css = @"
    body {
      font-family: Tahoma;
      font-size:10pt;
      background-color:white;
      color:#333;
    }

    .even {
      background-color:#eee;
    }
    
    .odd {
      background-color:#666;
    }

    .report-table {
      width: 100%;
    }
"@
  $date = Get-Date
  $reportTitle = "<h1>SharePoint Farm Report</h1><h2>$date</h2>"
  ConvertTo-EnhancedHTML -HTMLFragments $report -Title 'SharePoint Farm Report' -PreContent $reportTitle -CssStyleSheet $css -jQueryURI 'https://code.jquery.com/jquery-3.2.1.min.js' | Out-File -FilePath $ReportFilePath
}


 