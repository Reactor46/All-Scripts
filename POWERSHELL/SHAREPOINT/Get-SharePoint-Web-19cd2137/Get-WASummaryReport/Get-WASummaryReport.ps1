#------------------------------------------------------------------------------------------- 
# Name:            Get-WASummaryReport.ps1
# Description:     This script will get the Web Analytics Summary Report 
# Usage:        Run the function with the required parameters 
#                Context can be SPWebApplication, SPSite or SPWeb 
# By:             Ivan Josipovic, softlanding.ca 
#------------------------------------------------------------------------------------------- 
 
Function Get-WASummaryReport($Context,$DaysToGoBack){ 
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ea 0; 
    [System.Reflection.Assembly]::Load("Microsoft.Office.Server.WebAnalytics, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c") | Out-Null; 
    [System.Reflection.Assembly]::Load("Microsoft.Office.Server.WebAnalytics.UI, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c") | Out-Null; 
     
    Function DateTimeToDateId ([System.DateTime]$dt){ 
        if (![System.String]::IsNullOrEmpty($dt.ToString())){ 
            return [System.Int32]::Parse($dt.ToString("yyyyMMdd", [System.Globalization.CultureInfo]::InvariantCulture), [System.Globalization.CultureInfo]::InvariantCulture); 
        }else{ 
            return 0; 
        } 
    } 
     
    #Not used in this report but other report types require it. 
    Function GetSortOrder([String]$sortColumn,[Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.OrderType]$order){ 
        $SortOrders = New-Object System.Collections.Generic.List[Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.SortOrder]; 
        $sortOrders.Add((New-Object Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.SortOrder($sortColumn, $order))); 
        return ,$SortOrders 
    } 
     
    $AggregationContext = [Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.AggregationContext]::GetContext($Context); 
    if (!$?){throw "Cant get the Aggregation Context";} 
     
    $viewParamsList = New-Object System.Collections.Generic.List[Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.ViewParameterValue] 
    $viewParamsList.Add((New-Object Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.ViewParameterValue("PreviousStartDateId", (DateTimeToDateId([System.DateTime]::UtcNow.AddDays(-($DaysToGoBack * 2))))))); 
    $viewParamsList.Add((New-Object Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.ViewParameterValue("CurrentStartDateId", (DateTimeToDateId([System.DateTime]::UtcNow.AddDays(-($DaysToGoBack))))))); 
    $viewParamsList.Add((New-Object Microsoft.Office.Server.WebAnalytics.ProcessedDataRetriever.ViewParameterValue("Duration", $DaysToGoBack))); 
 
    $dataPacket = [Microsoft.Office.Server.WebAnalytics.Reporting.FrontEndDataRetriever]::QueryData($AggregationContext, $null, "fn_WA_GetSummary", $viewParamsList, $null, $null, 1, 25000, $False); 
    if (!$?){throw "Unable to get the Data. Try running the script as the Farm Account. If that doesnt work, make sure that the Web Analytics Service Application is connected to the Web Application and that the Site Web Analytics reports work through the browser.";} 
     
    return $dataPacket.DataTable 
} 
 
$WebApp = Get-SPWebApplication http://sp.client.com 
$Site = Get-SPSite http://sp.client.com 
$Web = Get-SPWeb http://sp.client.com 
 
Get-WASummaryReport -Context $WebApp -DaysToGoBack 30  
Get-WASummaryReport -Context $Site -DaysToGoBack 30 
Get-WASummaryReport -Context $Web -DaysToGoBack 30 