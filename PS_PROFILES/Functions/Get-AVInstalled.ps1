function Get-AntiVirusProduct { 
[CmdletBinding()] 
param ( 
[parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)] 
[Alias('name')] 
$computername=$env:computername 
)
$AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct  -ComputerName $computername

#Switch to determine the status of antivirus definitions and real-time protection. 
#The values in this switch-statement are retrieved from the following website: http://community.kaseya.com/resources/m/knowexch/1020.aspx 
switch ($AntiVirusProduct.productState) { 
"262144" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "262160" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "266240" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "266256" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "393216" {$defstatus = "Up to date" ;$rtstatus = "Disabled"} 
    "393232" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "393488" {$defstatus = "Out of date" ;$rtstatus = "Disabled"} 
    "397312" {$defstatus = "Up to date" ;$rtstatus = "Enabled"} 
    "397328" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
    "397584" {$defstatus = "Out of date" ;$rtstatus = "Enabled"} 
default {$defstatus = "Unknown" ;$rtstatus = "Unknown"} 
    }

#Create hash-table for each computer 
$ht = @{} 
$ht.Computername = $computername 
$ht.Name = $AntiVirusProduct.displayName 
$ht.ProductExecutable = $AntiVirusProduct.pathToSignedProductExe 
$ht.'Definition Status' = $defstatus 
$ht.'Real-time Protection Status' = $rtstatus

#Create a new object for each computer 
New-Object -TypeName PSObject -Property $ht

}