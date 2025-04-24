param (  
   [string]$DBserver = $(throw "Missing server name (please use -dbserver [dbserver])"),  
   [string]$path = $(throw "Missing input file (please use -path [path\file.txt])")  
 )  
 #Set Variables  
 $input = @(Get-Content $path)  
 #Addin SharePoint2010 PowerShell Snapin  
 Add-PSSnapin -Name Microsoft.SharePoint.PowerShell  
 #Declare Log File  
 Function StartTracing  
 {  
   $LogTime = Get-Date -Format yyyy-MM-dd_h-mm  
   $script:LogFile = "MissingSetupFileOutput-$LogTime.txt"  
   Start-Transcript -Path $LogFile -Force  
 }  
 #Declare SQL Query function  
 function Run-SQLQuery ($SqlServer, $SqlDatabase, $SqlQuery)  
 {  
   $SqlConnection = New-Object System.Data.SqlClient.SqlConnection  
   $SqlConnection.ConnectionString = "Server =" + $SqlServer + "; Database =" + $SqlDatabase + "; Integrated Security = True"  
   $SqlCmd = New-Object System.Data.SqlClient.SqlCommand  
   $SqlCmd.CommandText = $SqlQuery  
   $SqlCmd.Connection = $SqlConnection  
   $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter  
   $SqlAdapter.SelectCommand = $SqlCmd  
   $DataSet = New-Object System.Data.DataSet  
   $SqlAdapter.Fill($DataSet)  
   $SqlConnection.Close()  
   $DataSet.Tables[0]  
 }  
 #Declare the GetFileUrl function  
 function GetFileUrl ($filepath, $DBname)  
 {  
     #Define SQL Query and set in Variable  
     $Query = "SELECT * from AllDocs where SetupPath = '"+$filepath+"'"  
     #Runing SQL Query to get information about the MissingFiles and store it in a Table  
     $QueryReturn = @(Run-SQLQuery -SqlServer $DBserver -SqlDatabase $DBname -SqlQuery $Query | select Id, SiteId, DirName, LeafName, WebId, ListId)  
     foreach ($event in $QueryReturn)  
       {  
         if ($event.id -ne $null)  
         {  
         $site = Get-SPSite -Limit all | where { $_.Id -eq $event.SiteId }  
         #get the URL of the Web:  
         $web = $site | Get-SPWeb -Limit all | where { $_.Id -eq $event.WebId }  
         #Write the SPWeb URL to host  
         Write-Host $filepath -nonewline -foregroundcolor yellow  
         Write-Host ";" -nonewline  
         write-host $web.Url -NoNewline -foregroundcolor green  
         #get the URL of the actual file:  
         $file = $web.GetFile([Guid]$event.Id)  
         #Write the relative URL to host  
         write-host "/" -nonewline -foregroundcolor green  
         write-host $file.Url -foregroundcolor green 
         $file.delete()
         }  
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