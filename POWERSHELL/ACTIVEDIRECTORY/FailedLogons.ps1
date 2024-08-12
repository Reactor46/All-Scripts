$DateSave = get-date -format d.M.yyyy
 
$DC= "LASDC02.Contoso.corp" 
#$DC2 = "PHXDC03.phx.Contoso.corp"
#$DC3 = "LASAUTH01.creditoneapp.biz"

$Report= "C:\LazyWinAdmin\WinSysCheckList\Reports\Contoso.CORP\Failed-Logons-Report_$DateSave.html"
#$Report2= "C:\LazyWinAdmin\WinSysCheckList\Reports\PHX.Contoso.CORP\Failed-Logons-Report_$DateSave.html"
#$Report3= "C:\LazyWinAdmin\WinSysCheckList\Reports\CREDITONEAPP.BIZ\Failed-Logons-Report_$DateSave.html" 
 
$HTML=@" 
<title>Event Logs Report</title> 
 
BODY{background-color :#FFFFF} 
TABLE{Border-width:thin;border-style: solid;border-color:Black;border-collapse: collapse;} 
TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color: ThreeDShadow} 
TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color: Transparent} 
 
"@ 
 
$eventsDC= Get-Eventlog security -Computer $DC -InstanceId 4625 -After (Get-Date).AddDays(-7) | 
   Select TimeGenerated,ReplacementStrings | 
   % { 
     New-Object PSObject -Property @{ 
      Source_Computer = $_.ReplacementStrings[13] 
      UserName = $_.ReplacementStrings[5] 
      IP_Address = $_.ReplacementStrings[19] 
      Date = $_.TimeGenerated 
    } 
   } 
    
  $eventsDC | ConvertTo-Html -Property Source_Computer,UserName,IP_Address,Date -head $HTML -body "<H2>Gernerated On $DateSave</H2>"| 
     Out-File $Report -Append 