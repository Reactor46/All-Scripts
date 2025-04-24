#Changed script to use text file to load the webservers rather then type them all in. 5/3/17 Jim Adkins

$limit = (Get-Date).AddDays(-2)
$path = 'C:\inetpub\logs\LogFiles'
$Servers= Get-Content C:\LazyWinAdmin\WebServers\Configs\IISLogs.txt


function removeLogs
{
    Try {

            foreach ($Server in $Servers) {

                Write-Host $Server
                $path = '\\' + $Server +'\d$\LogFiles\IIS\'
                #$path = '\\' + $Server +'\C$\LogFiles\IIS\W3SVC2'
                $expression = 'Get-ChildItem -Path ' + $path + ' -Recurse -Force |
                Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force -ErrorAction Continue'
                
               # $expression = "Get-ChildItem -Path $path -Recurse -Force |
               # Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force -ErrorAction Continue"
                
                #Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force

  write-host $expression
                Invoke-Expression $expression  
  
            }#foreach 
     } #Try

       Catch [system.exception]
       {
            Write-Output "ERROR FOUND:"
            $error[0]
           # [Enviroment]::Exit("2")
       } #Catch
       Catch
       {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Output "There was an error:"
            $ErrorMessage
            $FailedItem
           # [Environment]::Exit("2")
       } #Catch


        Finally
        {
	        Write-Output "##############################################"
	        Write-Output "JOB COMPLETED SUCCESSFULLY"
	        Write-Output "##############################################"

foreach ($Server in $Servers){

    $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $Server -Filter "DeviceID='D:'" | Select-Object Size,FreeSpace
    $percentUsed = [math]::Round(100 - (($disk.FreeSpace / $disk.Size) * 100))
    Write-Host $Server "D:\ - % Used =>" $percentUsed % |fl
    $DiskUsed+=$Server + " D:\ - % Used =>"+ $percentUsed +"%`n"

}



            #Send-MailMessage -to "Nicholas.Fedei@creditone.com" -from "Task@Creditone.com" -subject "IIS Log Cleanup has Run" -Body $DiskUsed  -SmtpServer "lasexch01.Contoso.corp"

          #  [Environment]::Exit("1")
        } #Finally

} #Function
    


removeLogs