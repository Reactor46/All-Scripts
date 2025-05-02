
# ***********************************************************************
						# Variables initialization
# ***********************************************************************
$Temp = $env:temp
$ComputerName = gc env:computername
$All_System_Error = get-eventlog System   | where {$_.EntryType -eq "Error"} | select timegenerated, source, eventid, message
$All_Apps_Error = get-eventlog Application   | where {$_.EntryType -eq "Error"} | select timegenerated, source, eventid, message

$Date = get-date
$HTML_Events = $Temp + "\Events_List.html"
$CSS_File = $Temp + "\systanddeploy.css" # CSS for HTML Export

# *************************************************************************************************

$Title = "<p><span class=titre_list>Last applications and system errors on $ComputerName</span><br><span class=subtitle>This document has been updated on $Date</span></p><br>"


$System_Events = "<p class=New_object>Last 5 system errors</p>"	
$System_Events_b = $All_System_Error | select -first 5 | % { New-Object psobject -Property @{
Date= $_."timegenerated"	
Source=$_."source"
Event_ID = $_."eventid"	
Issue=$_."message"		
}}  | select Date, Source, Event_ID, Issue | ConvertTo-HTML -Fragment	

$System_Events = $System_Events + $System_Events_b

$Apps_Events = "<p class=New_object>Last 5 application errors</p>"	
$Apps_Events_b = $All_Apps_Error | select -first 5 | % { New-Object psobject -Property @{
Date= $_."timegenerated"	
Source=$_."source"
Event_ID = $_."eventid"	
Issue=$_."message"		
}}  | select Date, Source, Event_ID, Issue | ConvertTo-HTML -Fragment	

$Apps_Events = $Apps_Events + $Apps_Events_b


ConvertTo-HTML  -body " $Title<br>$System_Events<br><br>$Apps_Events" -CSSUri $CSS_File | 		
Out-File -encoding ASCII $HTML_Events
invoke-expression $HTML_Events

	


