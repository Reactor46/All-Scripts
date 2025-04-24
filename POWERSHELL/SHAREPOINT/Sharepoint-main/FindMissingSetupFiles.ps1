# Title:   FindMissingSetupFile.ps1
# Version: 1.0, 11 NOV 2021
# Author:  James Sanders
# Purpose: Find missing setup files to resolve Health Analyzer issues

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Function Run-SQLScript($SQLServer, $SQLDatabase, $SQLQuery) {
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

# Define configuration parameters
$Server="agl-sppre-d.ad.ing.net"
$Database="WSS_Content_Intranet"
$SetupFile="Features\CKS EBE Themes\Themes\Clean\images\Technorati.png"

# Query SQL Server content database to get information about the missing files
$Query = "SELECT * from AllDocs where SetupPath like '"+$SetupFile+"'"

$QueryResults = @(Run-SQLScript -SQLServer $Server -SQLDatabase $Database -SQLQuery $Query | select Id, SiteId, WebId)

# Iterate through results
ForEach ($Result in $QueryResults) {
  If ($Result.id -ne $Null) {
    $Site = Get-SPSite -Limit all | where { $_.Id -eq $Result.SiteId }
    $Web = $Site | Get-SPWeb -Limit all | where { $_.Id -eq $Result.WebId }
    #Get the URL of the file which is referring the feature
    $File = $web.GetFile([Guid]$Result.Id)
    write-host "$($Web.URL)/$($File.Url)" -foregroundcolor green


    $File.delete()
  }
}