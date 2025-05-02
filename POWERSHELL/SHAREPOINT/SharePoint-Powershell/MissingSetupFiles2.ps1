#Read more: https://www.sharepointdiary.com/2016/01/fix-missingsetupfile-error-test-spcontentdatabase-in-sharepoint-migration.html#ixzz7mFJlA9Rd
param (  
   [string]$DBServer = $(throw "Missing server name (please use -dbserver [dbserver])"),  
   [string]$path = $(throw "Missing input file (please use -path [path\file.txt])")  
 )  
#Set Variables  
 $input = @(Get-Content $path)  
 #Addin SharePoint2010 PowerShell Snapin  
 Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 #Declare Log File  
 Function StartTracing  
 {  
   $LogTime = Get-Date -Format yyyy-MM-dd_h-mm  
   $script:LogFile = "MissingSetupFileOutput-$LogTime.txt"  
   Start-Transcript -Path $LogFile -Force  
 }  
 #Declare SQL Query function  
Function Run-SQLScript($SQLServer, $SQLDatabase, $SQLQuery)
{
    $ConnectionString = "Server =" + $SQLServer + "; Database =" + $SQLDatabase + "; Integrated Security = True"
    $Connection = new-object system.data.SqlClient.SQLConnection($ConnectionString)
    $Command = new-object system.data.sqlclient.sqlcommand($SQLQuery,$Connection)
    $Connection.Open()
    $Adapter = New-Object System.Data.sqlclient.sqlDataAdapter $Command
    $Dataset = New-Object System.Data.DataSet
    $Adapter.Fill($Dataset)
    $Connection.Close()
    $Dataset.Tables[0]
}
 
#Define configuration parameters
$DBServer="Ab-SQL-001"
$Database="SP10_Content_KM"
$SetupFile="Features\Crescent.Intranet.WebParts\WEFnews\WEFnews.dwp"
 
#Query SQL Server content Database to get information about the MissingFiles
$Query = "SELECT * from AllDocs where SetupPath like '"+$SetupFile+"'"
$QueryResults = @(Run-SQLScript -SQLServer $Server -SQLDatabase $Database -SQLQuery $Query | select Id, SiteId, WebId)
 
#Iterate through results
foreach ($Result in $QueryResults)
{
    if($Result.id -ne $Null)
    {
        $Site = Get-SPSite -Limit all | where { $_.Id -eq $Result.SiteId }
        $Web = $Site | Get-SPWeb -Limit all | where { $_.Id -eq $Result.WebId }
 
        #Get the URL of the file which is referring the feature
        $File = $web.GetFile([Guid]$Result.Id)
        write-host "$($Web.URL)/$($File.Url)" -foregroundcolor green
         
        #$File.delete()
    }
}
#Start Logging  
 StartTracing  
 #Log the CVS Column Title Line  
 write-host "MissingSetupFile;Url" -foregroundcolor Yellow
 foreach ($event in $input)  
   {  
   $filepath = $event.split(";")[0]  
   $DBname = $event.split(";")[1]  
   #call Function  
   GetFileUrl $filepath $dbname  
   }  
 #Stop Logging  
 Stop-Transcript  

