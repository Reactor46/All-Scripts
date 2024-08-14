<#   #>
$VIServer = 'lasvcenter01.Contoso.corp'
$VMlist = @()

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
if(!($global:DefaultVIServer.IsConnected)){$VC_Session = Connect-VIServer $VIServer}


#Create list of VM's with less than 4 vCPU's
#$VMlist = Get-Cluster -Name LASESXAPP01-CIM | Get-VM | Where {$_.numCpu -le 3 -and  $_.PowerState -eq 'PoweredOn'}

    $ListFile = Get-Content C:\Temp\cpulist.txt #Use Myles List instead.
    $VMList =@()

ForEach($Obj in $ListFile){
        
 $VMList += (Get-VM -Name $Obj)

}


#$VMlist = Get-vm 'pstest02'

Function Add_vCPU {

        Param(
        [Array]$VM
    )


    #Shut down the VM.
    
 $Null = Shutdown-VMGuest $vm -Confirm:$False
 $VMstatus = Get-VM $vm
                
      While ($VMstatus.PowerState -eq 'PoweredOn'){
        Sleep  1
      Write-host '.' -NoNewline
      $VMstatus = Get-VM $vm}

      Write-host 'Shut down Complete'

    #Add the vCPU
    Set-VM -vm $vm -NumCpu 4 -CoresPerSocket 1 -Confirm:$false 

    #Start the VM 
    $Null = Start-vm -VM $vm.Name

}
foreach($vm in $VMlist){

    Write-Host ("Working with HostName {0}" -f $vm )
     Add_vCPU -VM $vm
    }

