function Get-Skew {
<#
      .SYNOPSIS
         Gets the time of a windows server
 
      .DESCRIPTION
         Uses WMI to get the time of a remote server
 
      .PARAMETER  ServerName
         The Server to get the date and time from
 
      .EXAMPLE
         PS C:\> Get-Skew -RemoteServer RemoteServer01 -LocalServer localhost
 
      
   #>


[CmdletBinding()]
   param(
      [Parameter(Position=0, Mandatory=$true)]
      [ValidateNotNullOrEmpty()]
      [System.String]
      $Servers 
       
   )

$RemoteServer = Get-Time -ServerName $Servers
$LocalServer = Get-Time -ServerName LASDC01
 
$Skew = $LocalServer.DateTime - $RemoteServer.DateTime
 
# Check if the time is over 30 seconds
If (($Skew.TotalSeconds -gt 30) -or ($Skew.TotalSeconds -lt -30)){
   Write-Host "Time is not within 30 seconds"
} Else {
   Write-Host "Time checked ok"
}
}