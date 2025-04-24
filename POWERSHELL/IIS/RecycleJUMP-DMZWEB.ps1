$Servers = GC .\DMZWEB.txt
ForEach($Srv in $Servers){
.\RecycleApplicationPool.ps1 -ApplicationPoolComputerName $Srv -WebAppPool "CreditOneBankJumphostAppPool" 
Get-AppPool -Server $Srv -Pool "CreditOneBankJumphostAppPool"
}