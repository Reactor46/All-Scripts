param([String] $servername = $(throw "Please specify the Servername"), [int32] $timerange = $(throw "Please specify a Time Range in Hours"),[String] $emailaddress= $(throw "Please Specify the Email address you wish to use"))
$dtQueryDT = [DateTime]::UtcNow.AddHours(-$timerange)
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($dtQueryDT)

$WmiNamespace = "ROOT\MicrosoftExchangev2"
$filter = "entrytype = '1020' and OriginationTime >= '" + $WmidtQueryDT + "' or entrytype = '1028' and OriginationTime >= '" + $WmidtQueryDT + "'"
get-wmiobject -class Exchange_MessageTrackingEntry -Namespace $WmiNamespace -ComputerName $servername -filter $filter | where-object  {$_.RecipientAddress -eq $emailaddress} | where-object {$_.SenderAddress -ne "<>"} | ft @{expression={[System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime)}; width=25; label="Time Sent"},SenderAddress,Subject



