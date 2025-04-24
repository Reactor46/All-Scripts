Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
 
Function WaitForInsallation([string] $SolutionName)
{
        Write-Host -NoNewline "Waiting for deployment job to complete" $SolutionName "."
        $WSPSol = Get-SPSolution $SolutionName
        while($wspSol.JobExists)
        {
            sleep 2
            Write-Host -NoNewline "."
            $wspSol = Get-SPSolution $SolutionName
        }
        Write-Host "job Completed" -ForegroundColor green
}
 
Function Deploy-SPSolution ($WSPFolderPath)
{
    #Get all wsp files from the given folder
    $WSPFiles = Get-childitem $WspFolderPath | where {$_.Name -like "*.wsp"}
 
    #Iterate through each wsp and Add in to the solution store
    ForEach($File in $wspFiles)
    {
        $wsp = Get-SPSolution | Where {$_.Name -eq $File.Name}
 
        if($wsp -eq $null)
        {
            write-host "Adding WSP solution:"$File.Name
            Add-SPSolution -LiteralPath ($WspFolderPath + "\" + $file.Name)
        }
        else
        {
            write-host "solution already exists!"
 
        }
    }
}
 
try
{
        Deploy-SPSolution "D:\Solutions"
}
catch
{
    write-host $_.exception
} 


#Read more: https://www.sharepointdiary.com/2015/04/deploy-multiple-solutions-in-sharepoint-using-powershell.html#ixzz7jbkoyiGG