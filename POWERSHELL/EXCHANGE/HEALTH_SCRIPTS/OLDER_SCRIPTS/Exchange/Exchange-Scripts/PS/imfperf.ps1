param([String] $servername = $(throw "Please specify the Servername"))
$WmiNamespace = "ROOT\CIMV2"
$filter = "Name='_total'"
$Qresults = get-wmiobject -class Win32_PerfRawData_MSExchangeUCF_MSExchangeIntelligentMessageFilter -Namespace $WmiNamespace -ComputerName $servername -filter $filter 
$format = @{Expression = {$_.TotalMessagesAssignedanSCLRatingof0};Label = "0"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof1};Label = "1"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof2};Label = "2"},
@{Expression = {$_.TotalMessagesAssignedanSCLRatingof3};Label = "3"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof4};Label = "4"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof5};Label = "5"},
@{Expression = {$_.TotalMessagesAssignedanSCLRatingof6};Label = "6"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof7};Label = "7"},@{Expression = {$_.TotalMessagesAssignedanSCLRatingof8};Label = "8"},
@{Expression = {$_.TotalMessagesAssignedanSCLRatingof9};Label = "9"},@{Expression = {$_.TotalUCEMessagesActedUpon};Label = "#Gate-Blk"},@{Expression = {$_.TotalMessagesScannedforUCE};Label = "Total"}
$Qresults | format-table $format

