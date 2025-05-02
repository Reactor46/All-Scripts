Function Get-IIS
{

$servers = (Get-Content c:\temp\servers.txt)

$IIS_S= ""| Select VM,Status,Version
 
foreach($vm in $servers){


$iis = get-wmiobject Win32_Service -ComputerName $vm -Filter "name='IISADMIN'"
 
if($iis.State -eq "Running")
{
$keyname = 'SOFTWARE\\Microsoft\\InetStp' 
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $vm) 
$key = $reg.OpenSubkey($keyname) 
$iisInfo = $key.GetValue('VersionString') 


$IIS_S.Version= $iisInfo
$IIS_S.Status= "Running"
$IIS_S.VM =$vm
}
else
{
$IIS_S.Version = ""
$IIS_S.Status= "Not Running"
$IIS_S.VM =$vm

}
}

$IIS_S |epcsv C:\temp\IIS.csv

}

