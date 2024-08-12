##############################################################################
##                                                                                            
##           Author: Vikas Sukhija                  		                              
##           Date: 03/03/2015                       		      			      
##           Description:- Add particular user/group to Localadmin of multiple servers  
##              			      		                              			      
##############################################################################

$servers = import-csv .\localadmin.csv

$domain = "domain"
foreach($i in $servers){

$server= $i.server
$usgroup = $i.usgroup

Write-host "Adding $usgroup to server $server" -foregroundcolor green

$User = [ADSI]("WinNT://$domain/$usgroup")
$Group = [ADSI]("WinNT://$server/Administrators")
$Group.PSBase.Invoke("Add",$User.PSBase.Path)

}
##############################################################################