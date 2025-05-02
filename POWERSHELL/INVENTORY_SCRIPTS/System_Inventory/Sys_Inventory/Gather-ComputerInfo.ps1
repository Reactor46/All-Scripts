<#
.Synopsis
   Gather Computer Inventory Information from your network using AD
.DESCRIPTION
   This function will gather information from your local computer network, using get-adcomputer to
   retrieve the list of computer names to work on, then it will use GET-WMIOBJECT (Whammy object)
   commands to get detailed information: Hard Drive Size/Freespace, RAM Size, MAC Address, information
   about IE, JAVA and Office, and other miscellaneous info like Computer Manufacturer, Model, Username
   (logged on at the time of scan...).  I personally use the location field in AD to store asset tag #s
   that we label our computers with, you can change this, or delete that parameter out of 'select' in
   that particular command....up to you.  Have fun.
.EXAMPLE
    After changing to the directory you would like to store your daily inventory, run this command:
    Gather-ComputerInfo -SaveDirectory .\
.EXAMPLE
    Gather-ComputerInfo -SaveDirectory C:\Users\jcummings\Documents\Inventory
#>
function Gather-ComputerInfo
{
    [CmdletBinding()]
    Param
    (
        # Only one parameter: The directory where the information will be stored....maybe a temp file location
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $savedirectory
    )

    Begin
    {
        # This stores the basic information for all computers in file based on date...
        Get-ADComputer -filter * -Property * | Select-Object Name,Description,OperatingSystem,OperatingSystemServicePack,IPv4Address,Enabled,Location | Export-Csv ("$savedirectory\adcomp" + (get-date -format "MM_dd_yyyy") + ".csv") -NoTypeInformation -Encoding UTF8
        $computers = import-csv ("$savedirectory\adcomp" + (get-date -format "MM_dd_yyyy") + ".csv")
    }
    Process
    {
        foreach ( $c in $computers.name )
        {
            if (  Test-Connection -ComputerName $c -BufferSize 16 -Count 1 -Quiet )
            {

                "Scanning: $c" # | Out-File ("$savedirectory\adcomp" + (get-date -format "MM_dd_yyyy") + ".txt") -append
            
                <# Hard Drive Info #>
                get-WmiObject win32_logicaldisk -Computername $c | select PSComputername,Freespace,Size | Export-Csv ("$savedirectory\ADDrive" + (get-date -format "MM_dd_yyyy") + ".csv") -Append -NoTypeInformation -Encoding UTF8

                <# Ram, Manufacturer, Model, System Type(64 or 32bit), Admin logon and User logged on during scan #>
                Get-WmiObject win32_computersystem -ComputerName $c | Select-Object PSComputerName,Manufacturer,Model,PrimaryOwnerName,SystemType,TotalPhysicalMemory,UserName | Export-Csv ("$savedirectory\ADMisc" + (get-date -format "MM_dd_yyyy") + ".csv") -Append -NoTypeInformation -Encoding UTF8

                <# IE Version #>
                $filename = "\Program Files\Internet Explorer\iexplore.exe"
                $obj = New-Object System.Collections.ArrayList
                $filepath = Test-Path "\\$c\c$\$filename"
                if ($filepath -eq "True")
                {
                    $file = Get-Item "\\$c\c$\$filename"
                    $obj += New-Object psObject -Property @{'Computer'=$c;'IE Version'=$file.VersionInfo|Select-Object FileVersion;'Length'=$file.Length}
                    $obj | Export-Csv ("$savedirectory\ADIEVer" + (get-date -format "MM_dd_yyyy") + ".csv") -append -notypeinformation -encoding UTF8
                }

                <# JAVA Version #>
                Get-WmiObject win32_installedwin32program -ComputerName $c | Where-Object { $_.vendor -eq "Oracle" } | Select-Object PSComputername,Vendor,Version | Export-Csv ("$savedirectory\ADJAVA" + (get-date -format "MM_dd_yyyy") + ".csv") -Append -NoTypeInformation -Encoding UTF8

                <# Office Version 15.0.4569.1506 = Office 2013 #>
                <# Note: The Star*** is necessary in this next line #>
                Get-WmiObject win32_installedwin32program -ComputerName $c | Where-Object { $_.Name -like "Microsoft Outlook*" -and $_.vendor -eq "Microsoft Corporation" -and $_.version -eq "15.0.4569.1506" } | Select-Object PSComputername,Vendor,Version | Export-Csv ("$savedirectory\ADOffice" + (get-date -format "MM_dd_yyyy") + ".csv") -Append -NoTypeInformation -Encoding UTF8

                <# MAC Address #>
                Get-WmiObject win32_networkadapterconfiguration -ComputerName $c | Where-Object { $_.defaultipgateway -eq "192.168.1.240" } | Select-Object MacAddress,PSComputerName |Export-Csv ("$savedirectory\ADMAC" + (get-date -format "MM_dd_yyyy") + ".csv") -Append -NoTypeInformation -Encoding UTF8
            } else {
                "Skipping: $c"
            }
        }
    }
    End
    {
        <# Load in file with most records... #>
        $a = Import-Csv ("$savedirectory\adcomp" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Hard Drive Info, without output adjusted to show in GB and only 2 decimal places. #>
        $b = Import-Csv ("$savedirectory\ADDrive" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "HD Size" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "HD Free" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( ($b[$j].pscomputername -eq $a[$i].Name) -and ($a[$i]."HD Size" -eq "null"))
              {
                 <# So update A with B's information #>
                 <# Note adjustment to show result with 2 decimal places, and converted to GB #>
                 $a[$i]."HD Free" = ("{0:N2}" -f ($b[$j].Freespace / 1GB))
                 $a[$i]."HD Size" = ("{0:N2}" -f ($b[$j].Size / 1GB))
              }
           }
        }

        <# RAM and computer type information... #>
        $b = Import-Csv ("$savedirectory\ADMisc" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "Manufacturer" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "Model" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "PrimaryOwnerName" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "SystemType" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "TotalPhysicalMemory" -NotePropertyValue "null"
        $a | Add-Member -NotePropertyName "UserName" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( $b[$j].PSComputername -eq $a[$i].Name)
              {
                 <# So update A with B's information #>
                 $a[$i]."Manufacturer" = $b[$j]."Manufacturer"
                 $a[$i]."Model" = $b[$j]."Model"
                 $a[$i]."PrimaryOwnerName" = $b[$j]."PrimaryOwnerName"
                 $a[$i]."SystemType" = $b[$j]."SystemType"
                 $a[$i]."TotalPhysicalMemory" = ("{0:N2}" -f ($b[$j].TotalPhysicalMemory / 1GB))
                 $a[$i]."UserName" = $b[$j]."UserName"

              }
           }
        }

        <# IE Version... #>
        $b = Import-Csv ("$savedirectory\ADIEVer" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "IE Version" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( $b[$j].computer -eq $a[$i].Name)
              {
                 <# So update A with B's information #>
                 <# NOTICE the extra bit at the end here for extracting the IE Version number... #>
                 <# Thanks to http://www.lazywinadmin.com/2013/10/powershell-get-substring-out-of-string.html #>
                 $a[$i]."IE Version" = $b[$j]."IE Version".Split(".")[0].substring(14)
              }
           }
        }

        <# JAVA Version... #>
        $b = Import-Csv ("$savedirectory\ADJAVA" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "JAVA" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( $b[$j].PScomputername -eq $a[$i].Name)
              {
                 <# So update A with B's information #>
                 $a[$i]."JAVA" = $b[$j].Version
              }
           }
        }

        <# Office Version... #>
        $b = Import-Csv ("$savedirectory\ADOffice" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "Office" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( $b[$j].pscomputername -eq $a[$i].Name)
              {
                 <# So update A with B's information #>
                 if ( $b[$j].Version -eq "15.0.4569.1506" )
                 {
                    $a[$i]."Office" = "2013"
                 }
                if ( $b[$j].Version -eq "16.0.4266.1001" )
                 {
                    $a[$i]."Office" = "2016"
                 }
               }
           }
        }

        <# MAC Address... #>
        $b = Import-Csv ("$savedirectory\ADMAC" + (get-date -format "MM_dd_yyyy") + ".csv")

        <# Add new column in A for value in B #>
        $a | Add-Member -NotePropertyName "MAC Addr" -NotePropertyValue "null"

        <# This is equivalent to the Excel Vlookup function with a range #>
        <# Cycle through all computers in B #>
        for ( $j=0; $j -lt $b.count-1; $j++ )
        {
           <# Find computer in A to set value from B #>
           for ( $i=0; $i -lt $a.count-1; $i++ )
           {
              <# We found the match here... #>
              if ( $b[$j].pscomputername -eq $a[$i].Name)
              {
                 <# So update A with B's information #>
                 $a[$i]."MAC Addr" = $b[$j].MacAddress
              }
           }
        }

        <# Save Updated File as CSV #>
        $a | Export-Csv ("$savedirectory\adnext" + (get-date -format "MM_dd_yyyy") + ".csv") -NoTypeInformation -Encoding UTF8
    }
}