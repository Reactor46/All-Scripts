function Get-SPRDiagnosticLoggingSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPDiagConfig
  )



 
    $properties = [ordered]@{
        'AllowLegacyTraceProviders' = $SPDiagConfig.AllowLegacyTraceProviders
        'AppAnalyticsAutomaticUploadEnabled' = $SPDiagConfig.AppAnalyticsAutomaticUploadEnabled
        'CustomerExperienceImprovementProgramEnabled' = $SPDiagConfig.CustomerExperienceImprovementProgramEnabled
        'DaysToKeepLogs' = $SPDiagConfig.DaysToKeepLogs
        'DownloadErrorReportingUpdatesEnabled' = $SPDiagConfig.DownloadErrorReportingUpdatesEnabled
        'ErrorReportingAutomaticUploadEnabled' = $SPDiagConfig.ErrorReportingAutomaticUploadEnabled
        'ErrorReportingEnabled' = $SPDiagConfig.ErrorReportingEnabled 
        'EventLogFloodProtectionEnabled' = $SPDiagConfig.EventLogFloodProtectionEnabled
        'EventLogFloodProtectionNotifyInterval' = $SPDiagConfig.EventLogFloodProtectionNotifyInterval
        'EventLogFloodProtectionQuietPeriod' = $SPDiagConfig.EventLogFloodProtectionQuietPeriod
        'EventLogFloodProtectionThreshold' = $SPDiagConfig.EventLogFloodProtectionThreshold
        'EventLogFloodProtectionTriggerPeriod' = $SPDiagConfig.EventLogFloodProtectionTriggerPeriod
        'LogCutInterval' = $SPDiagConfig.LogCutInterval
        'LogDiskSpaceUsageGB' = $SPDiagConfig.LogDiskSpaceUsageGB
        'LogLocation' = $SPDiagConfig.LogLocation
        'LogMaxDiskSpaceUsageEnabled' = $SPDiagConfig.LogMaxDiskSpaceUsageEnabled
        'ScriptErrorReportingDelay' = $SPDiagConfig.ScriptErrorReportingDelay
        'ScriptErrorReportingEnabled' = $SPDiagConfig.ScriptErrorReportingEnabled
        'ScriptErrorReportingRequireAuth' =$SPDiagConfig.ScriptErrorReportingRequireAuth
    }
    


    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  
}