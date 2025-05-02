##############################################################################
##                                                                                            
##           Author: Vikas Sukhija                  		                              
##           Date: 03/03/2015                       		      			      
##           Description:- Remove particular user/group from Localadmin of multiple servers  
##              			      		                              			      
##############################################################################

$servers = import-csv .\localadmin.csv

$domain = "domain"
foreach($i in $servers){

$server= $i.server
$usgroup = $i.usgroup

Write-host "Removing $usgroup to server $server" -foregroundcolor green

$User = [ADSI]("WinNT://$domain/$usgroup")
$Group = [ADSI]("WinNT://$server/Administrators")
$Group.PSBase.Invoke("Remove",$User.PSBase.Path)

}
##############################################################################