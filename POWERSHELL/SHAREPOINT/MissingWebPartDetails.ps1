#Read more: https://www.sharepointdiary.com/2016/01/fix-missingwebpart-issue-in-sharepoint-test-spcontentdatabase.html#ixzz7mFhYvMFO
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
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
$Database="WSS_Content_KM"
$webPartID="90F70A36-4EC6-167D-792C-31B7EA83201B"
 
#Get the Database Server from Database
$server = (Get-SPDatabase |?{ $_.name -eq $Database}).server
 
#Query SQL Server content Database to get information about the MissingFiles
$Query = "SELECT distinct ID,SiteId,DirName, LeafName, WebId, ListId from dbo.AllDocs where id in (select tp_PageUrlID from dbo.AllWebParts where tp_WebPartTypeId ='$($webPartID)')"
$QueryResults = Run-SQLScript -SQLServer $Server -SQLDatabase $Database -SQLQuery $Query | select Id , SiteId, DirName, LeafName, WebId, ListId
 
#Iterate through results
foreach ($Result in $QueryResults)
{
    if($Result.id -ne $Null)
    {
        $Site = Get-SPSite -Limit all | where { $_.Id -eq $Result.SiteId }
        $Web = $Site | Get-SPWeb -Limit all | where { $_.Id -eq $Result.WebId }
 
        #Get the URL of the file which is referring the web part
        $File = $web.GetFile([Guid]$Result.Id)
        write-host "$($Web.URL)/$($File.Url)" -foregroundcolor green
    }
}


