#Get all details of the error
Get-SPLogEvent | ?{$_.Correlation -eq "10687ea0-2bbe-30d3-cef3-2a557f71daab"}
 
#Retrieve selective columns of the error
Get-SPLogEvent | ?{$_.Correlation -eq "f53b559c-f70e-002f-694c-7d3b8b55f534"} | select Area, Category, Level, EventID, Message | Format-List
 
#Get sharepoint 2013 correlation id in error messages and send to file
Get-SPLogEvent | ? {$_Correlation -eq "3410f29b-b756-002f-694c-7a574ff74cab" } | select Area, Category, Level, EventID, Message | Format-List &gt; C:\SPError.log
 
#Get all issues logged in the past 10 minutes
Get-SPLogEvent -starttime (Get-Date).AddMinutes(-10)
 
#Get Events between specific time frames
Get-SPLogEvent -StartTime "03/06/2015 18:00" -EndTime "03/06/2015 18:30"


#Read more: https://www.sharepointdiary.com/2013/02/sharepoint-2013-correlation-id-get-detailed-error-using-powershell.html#ixzz7mLKhgngD