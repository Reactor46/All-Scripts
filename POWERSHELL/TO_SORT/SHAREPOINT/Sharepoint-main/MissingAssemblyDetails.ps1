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
    $script:LogFile = "MissingAssemblyOutput-$LogTime.csv"
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


function GetAssemblyDetails ($assembly, $DBname)
    {

    #Define SQL Query and set in Variable
    $Query = "SELECT * from EventReceivers where Assembly = '"+$assembly+"'"
    #$Query = "SELECT * from EventReceivers where Assembly = 'Microsoft.Office.InfoPath.Server, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c'" 

    #Runing SQL Query to get information about Assembly (looking in EventReceiver Table) and store it in a Table
    $QueryReturn = @(Run-SQLQuery -SqlServer $DBserver -SqlDatabase $DBname -SqlQuery $Query | select Id, Name, SiteId, WebId, HostId, HostType)

    #Actions for each element in the table returned
        foreach ($event in $QueryReturn)
        {   
            #HostID (check http://msdn.microsoft.com/en-us/library/ee394866(v=prot.13).aspx for HostID Type reference)
            if ($event.HostType -eq 0)
             {
             $site = Get-SPSite -Limit all | where {$_.Id -eq $event.SiteId}
             #Get the EventReceiver Site Object
             $er = $site.EventReceivers | where {$_.Id -eq $event.Id}
             
             Write-Host $assembly -nonewline -foregroundcolor yellow
             write-host ";" -nonewline
             write-host $site.Url -nonewline -foregroundcolor gray
             write-host ";" -nonewline
             write-host $er.Name -foregroundcolor green -nonewline
             write-host ";" -nonewline
             write-host $er.Class -foregroundcolor cyan
             #$er.Delete()
             }
             
             if ($event.HostType -eq 1)
             {
             $site = Get-SPSite -Limit all | where {$_.Id -eq $event.SiteId}
             $web = $site | Get-SPWeb -Limit all | where {$_.Id -eq $event.WebId}
             #Get the EventReceiver Site Object
             $er = $web.EventReceivers | where {$_.Id -eq $event.Id}
             $er.Name
             
             Write-Host $assembly -nonewline -foregroundcolor yellow
             write-host ";" -nonewline
             write-host $web.Url -nonewline -foregroundcolor gray
             write-host ";" -nonewline
             write-host $er.Name -foregroundcolor green -nonewline
             write-host ";" -nonewline
             write-host $er.Class -foregroundcolor cyan
             #$er.Delete()
             }
             
             if ($event.HostType -eq 2)
             {
             $site = Get-SPSite -Limit all | where {$_.Id -eq $event.SiteId}
             $web = $site | Get-SPWeb -Limit all | where {$_.Id -eq $event.WebId}
             $list = $web.Lists | where {$_.Id -eq $event.HostId}
             #Get the EventReceiver List Object
             $er = $list.EventReceivers | where {$_.Id -eq $event.Id}
             
             Write-Host $assembly -nonewline -foregroundcolor yellow
             write-host ";" -nonewline
             write-host $web.Url -nonewline -foregroundcolor gray
             write-host "/" -nonewline -foregroundcolor gray
             write-host $list.RootFolder -nonewline -foregroundcolor gray
             write-host ";" -nonewline
             write-host $er.Name -foregroundcolor green -nonewline
             write-host ";" -nonewline
             write-host $er.Class -foregroundcolor cyan
             #$er.Delete()
             }
              
        }
    }

#Start Logging
StartTracing

#Log the CVS Column Title Line
write-host "Assembly;Url;EventReceiverName;EventReceiverClass" -foregroundcolor Red

foreach ($event in $input)
    {
    $assembly = $event.split(";")[0]
    $DBname = $event.split(";")[1]
    GetAssemblyDetails $assembly $dbname
    }
    
#Stop Logging
Stop-Transcript

