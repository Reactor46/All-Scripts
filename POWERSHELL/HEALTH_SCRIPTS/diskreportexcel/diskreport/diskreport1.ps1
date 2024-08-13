########Working with Excel and Powershell to Generate Disk Space Report of Servers###########
#                   Author  : Abhishek Gupta                                                #
#                   Created : 2/22/2016                                                     #
############# Get the list of Servers########################################################
$servername = Get-Content ".\Servername.txt"
############Create excel COM object################
$excel = New-Object -ComObject excel.application
############Add a workbook#######################
$workbook = $excel.Workbooks.Add()
#############Remove other worksheets############
1..2 | ForEach {
    $Workbook.worksheets.item(2).Delete()
}
############# Connect to worksheet & Name the worksheet as DriveSpace###### 
$diskSpacewksht= $workbook.Worksheets.Item(1)
$diskSpacewksht.Name = 'DriveSpace'
######Create a header and make it bold and add color#########
$diskSpacewksht.Cells.Item(1,1) = 'ServerName'
$diskSpacewksht.Cells.Item(1,1).Font.Bold=$True
$diskSpacewksht.Cells.Item(1,1).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,2) = 'DeviceID'
$diskSpacewksht.Cells.Item(1,2).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,2).Font.Bold=$True
$diskSpacewksht.Cells.Item(1,3) = 'VolumeName'
$diskSpacewksht.Cells.Item(1,3).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,3).Font.Bold=$True
$diskSpacewksht.Cells.Item(1,4) = 'Size(GB)'
$diskSpacewksht.Cells.Item(1,4).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,4).Font.Bold=$True
$diskSpacewksht.Cells.Item(1,5) = 'FreeSpace(GB)'
$diskSpacewksht.Cells.Item(1,5).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,5).Font.Bold=$True
$diskSpacewksht.Cells.Item(1,6) = 'FreeSpacePercentage'
$diskSpacewksht.Cells.Item(1,6).Interior.ColorIndex = 24
$diskSpacewksht.Cells.Item(1,6).Font.Bold=$True
$row = 2
$column = 1
######Looping to get details each Server #########
foreach ($Server in $servername)
{
###### Code to get Disk details from server#####
$diskresult = Get-WmiObject win32_logicaldisk -ComputerName $Server  | Where-Object { $_.DriveType -eq 3 } 
ForEach ($disk in $diskresult) {
    #ServerName
    $diskSpacewksht.Cells.Item($row,$column) = $Server
    $column++
    #DeviceID
    $diskSpacewksht.Cells.Item($row,$column) = $disk.DeviceID
    $column++
    #VolumeName
    $diskSpacewksht.Cells.Item($row,$column) = $disk.VolumeName
    $column++
    #Size
    $diskSpacewksht.Cells.Item($row,$column) = "{0:N1}" -f($disk.Size /1GB)
    $column++
    #FreeSpace
    $diskSpacewksht.Cells.Item($row,$column) = "{0:N1}" -f($disk.FreeSpace /1GB)
    $column++
    #FreeSpacePercentage
    $criteria = ($disk.FreeSpace / $disk.Size)*100
    #$criteria = $diskSpacewksht.Cells.Item($row,$column) = ("{0:P}" -f ($disk.FreeSpace / $disk.Size))
    ######## if free space is less than 5 %#### Mark the cell red ######
    if ($criteria -lt '5')
    {
    $diskSpacewksht.Cells.Item($row,$column) = $criteria 
    $diskSpacewksht.Cells.Item($row,$column).Interior.ColorIndex = 3
    }
    else
    {
    $diskSpacewksht.Cells.Item($row,$column) = $criteria 
    $diskSpacewksht.Cells.Item($row,$column).Interior.ColorIndex = 43
    }
    #Increment to next Row and reset Column
    $row++
    $column = 1
}
}
#####Auto-sizing of Cells to fit the details ####
$usedRange = $diskSpacewksht.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null 
####### Delecting old output file on the location##############################
$output = "C:\diskreport.xlsx" 
$checkrep = Test-Path $output   
If ($checkrep -like "True")  
{  
Remove-Item $output  
}  
######## Save the Excel ##############
$workbook.SaveAs("C:\diskreport.xlsx")
$excel.Quit()
 
##########Parmaters to Send Mail########################################################   
$messageParameters = @{                            
                Subject = "Generate Disk Space Report of Servers"                            
                Body = "Generate Disk Space Report of Servers"                      
                From = "abhishek.gupta@lab.com"                            
                To = "abhishek.gupta3@lab.com" 
                Attachments = $output                   
                SmtpServer = "smtp relay server" 
            }    
Send-MailMessage @messageParameters    
##########Stop Transcript############################################  