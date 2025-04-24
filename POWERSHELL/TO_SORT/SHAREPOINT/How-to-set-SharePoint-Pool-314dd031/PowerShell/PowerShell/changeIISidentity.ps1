<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#> 

#import  WebAdministration module
Import-Module -Name "Webadministration"
$AppCol = @()
#Get IIS pool
$Col = Get-ChildItem -Path IIS:\AppPools
$PoolID = 0
#Loop the pools 
Foreach($item in $Col)
{
    If($item.processModel.identityType -eq "SpecificUser")
    {
        $obj = New-Object PSobject -Property @{ID= $PoolID+1;Name=$item.Name;State = $item.state;Identity=$item.processModel.userName}
        $AppCol += $obj
        $poolid++
    }
}
Write-Host "Find the following AppPools with specific user:"
$AppCol | select name,state,identity,id
Try
{
    #Get credential
    $Choice = Read-Host "Input the apppools to set identity(Default is all pools displayed)"
    $Cre = Get-Credential
    $userName = $Cre.UserName
    $SecurePassword = $Cre.Password
    #Convert secure string to plain text
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR) 

    If($Choice)
    {
    
        $NumCol = $Choice -split ","
        $NumCol
        Foreach($Num in $NumCol)
        {
           If($Num -lt $AppCol.Count)
           {
           
                $item =  $AppCol[$Num-1]
                $objPool = Get-ChildItem -Path IIS:\AppPools | Where-Object {$_.Name -eq $item.Name}
                $objpool.processModel.userName = $userName
                $objpool.processModel.password =  $PlainPassword
                $objPool | Set-Item    
                Write-Host "Set Applicatuin Pool identity "  $item.name  " successfully." -ForegroundColor Green
            

           }
           else
           {
                Write-Warning "The specific number is not in the array."
           }
        }
    }
    ElseIf($Choice -eq "")
    {
        Foreach($item in $AppCol)
        {
        
            $objPool = Get-ChildItem -Path IIS:\AppPools | Where-Object {$_.Name -eq $item.Name}
            $objpool.processModel.userName = $userName
            $objpool.processModel.password =  $PlainPassword
            $objPool | Set-Item    
            Write-Host "Set Applicatuin Pool identity "  $item.name  " successfully." -ForegroundColor Green
       

        }
    }
    Else
    {
        Write-Error "Invalid input value."
    }
}
Catch
{
    write-host $_ -ForegroundColor Red
}
cmd /c pause