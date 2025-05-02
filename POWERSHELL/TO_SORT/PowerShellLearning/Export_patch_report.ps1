#Defining a file path variable
$filePath = 'C:\Users\silad\Documents\report.csv'

#Defining a variable to store available hotfixes in. To get this info, use the Get-Hotfix command.
$hotfixes = Get-Hotfix

#Piping the output of the above command to a command that will export this as a csv report.
$hotfixes | Export-Csv $filePath -NoTypeInformation

#If statement to check if the csv file has been screated successfully
if (Test-Path $filePath){
Write-Host "Hotfix report exported to $filePath"
}else {
Write-Host "Failed to export report"
}