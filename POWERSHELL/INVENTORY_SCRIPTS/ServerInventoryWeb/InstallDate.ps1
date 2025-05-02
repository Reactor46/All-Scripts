<# 
    Author          : Preenesh Nayanasudhan 
    Script          : Get Install Date and OS Version for list of Servers provided in Text file Serverlist.txt 
    Purpose         : Get Install Date and OS Version for list of Servers provided in Text file and publish the result in CSV 
    Pre-requisite   : Create a Text File Serverlist.txt in the same path you save this script 
#> 
 
# Input File

$Servers = Get-Content .\Configs\Contoso.txt

$report = @()

ForEach ($server in $Servers) 
{ 
    #Test if computer is online
     
    if (test-Connection -ComputerName $server -Count 3 -Quiet ) 
    { 
        $PingResult = 'Server IS Pinging'
        
        # Connect to WMI on remote machine to get the required information

        $gwmios = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server

        # Get the Install Date and Convert readable date & time format

        $osidate = $gwmios.ConvertToDateTime($gwmios.InstallDate)
        $osname = $gwmios.Caption 

        $tempreport = New-Object PSObject 
        $tempreport | Add-Member NoteProperty 'Server Name' $server
        $tempreport | Add-Member NoteProperty 'Ping Result' $PingResult
        $tempreport | Add-Member NoteProperty 'OS Version' $osname 
        $tempreport | Add-Member NoteProperty 'OS Install Date' $osidate
        $report += $tempreport 
    }  
    else  
        {  
        $PingResult = 'Server NOT Pinging' 
        $tempreport = New-Object PSObject 
        $tempreport | Add-Member NoteProperty 'Server Name' $server
        $tempreport | Add-Member NoteProperty 'Ping Result' $PingResult
        $report += $tempreport 
    } 
 
} 
$report | Export-Csv -NoTypeInformation ('.\OSInstallDateContoso.csv')