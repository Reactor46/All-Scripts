function global:Test-Online {
#Requires -Version 2.0            
[CmdletBinding()]            
 Param             
   (                       
    [Parameter(Mandatory=$true,
               Position=0,                          
               ValueFromPipeline=$true,            
               ValueFromPipelineByPropertyName=$true)]            
    [Object]$Object,
    [String]$Parameter
    
   )#End Param

Begin            
{            
 Write-Verbose "`n Testing Server Response . . . "
 $i = 0            
}#Begin          
Process            
{
    $Object | ForEach-Object {
        if ($Parameter)
            {
                $Computer = $_."$Parameter"
            }
        elseif ($_.ComputerName)
            {
                $Computer = $_.ComputerName    
            }
        else    
            {
                $Computer = $_
            }
        
        if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -EA 0)
            {
                $i++
                Write-Verbose "Online : $($Computer)"
                $_
            }
         else
            {
                Write-Host "Offline: $($Computer)" -background red -foreground white
            }
    }#Foreach-Object (Computers)
    
}#Process
End
{
    "`n$($i) Computers online." | Out-Host
}#End

}#Test-Online 

