# ============================================================================================== 
# NAME: Listing Administrators and PowerUsers on remote machines  
#  
# AUTHOR: Mohamed Garrana ,  
# DATE  : 09/04/2010 
#  
# COMMENT:  
# This script runs against an input file of computer names , connects to each computer and gets a list of the users in the  local Administrators  
#and powerusers Groups . the output can be a csv file which can be readable on excel with all the computers from the input file 
# ============================================================================================== 
function Get-LocalServerAdmins { 
        param( 
    [Parameter(Mandatory=$true,valuefrompipeline=$true)] 
    [string]$strComputer) 
    begin {} 
    Process { 
        $adminlist ="" 
        #$powerlist ="" 
        $computer = [ADSI]("WinNT://" + $strComputer + ",computer") 
        $AdminGroup = $computer.psbase.children.find("Administrators") 
        #$powerGroup = $computer.psbase.children.find("Power Users") 
        $Adminmembers= $AdminGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        #$Powermembers= $PowerGroup.psbase.invoke("Members") | %{$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)} 
        foreach ($admin in $Adminmembers) { $adminlist = $adminlist + $admin + "," } 
        #foreach ($poweruser in $Powermembers) { $powerlist = $powerlist + $poweruser + "," } 
        $Computer = New-Object psobject 
        $computer | Add-Member noteproperty ComputerName $strComputer 
        $computer | Add-Member noteproperty Administrators $adminlist 
        #$computer | Add-Member noteproperty PowerUsers $powerlist 
        Write-Output $computer 
 
 
        } 
end {} 
} 
 
#Get-Content C:\LazyWinAdmin\Servers\RESULTS\Alive\All.txt | Get-LocalServerAdmins | Export-Csv 'C:\LazyWinAdmin\Local Admin Accounts\LocalServerAdmins.csv' -NoTypeInformation 