<#                                Export All AD Objects Powershell 

     *********************************** By allenage.com********************************************************
                                    

#>
$path = "C:\LazyWinAdmin\ActiveDirectory\AD_Audit\All Objects"
#ad computers
Get-ADComputer -Filter * -Properties ipv4Address, OperatingSystem,LastLogonDate,description  |
    Select-Object name, ipv4*, OperatingSystem,LastLogonDate,description |
    Sort LastLogondate -Descending  |
    Export-csv $path\ADComputers.csv -nti


#users

Import-module activedirectory
                  $data=@(
                  Get-ADUser  -filter * -Properties * |  
                  Select-Object @{Label = "First Name";Expression = {$_.GivenName}},  
                  @{Name = "Last Name";Expression = {$_.Surname}},
                  @{Name = "Full address";Expression = {$_.StreetAddress}}, 
                  @{Name = "City";Expression = {$_.City}}, 
                  @{Name = "State";Expression = {$_.st}}, 
                  @{Name = "Post Code";Expression = {$_.PostalCode}}, 
                  @{Name = "Country/Region";Expression ={$_.Country}},
                  @{Name = "MobileNumber";Expression = {$_.mobile}},
                  @{Name = "Phone";Expression = {$_.telephoneNumber}}, 
                  @{Name = "Description";Expression = {$_.Description}},
                  @{name=  "OU";expression={$_.DistinguishedName.split(',')[1].split('=')[1]}},
                  @{Name = "Email";Expression = {$_.Mail}},
                  @{Name = "Group Memberships";e= { ( $_.memberof | % { (Get-ADObject $_).Name }) -join “,” }},
                  @{Name = "UserPrincipalName";Expression = {$_.UserPrincipalName}},
                  @{Name = "Account Status";Expression = {if (($_.Enabled -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}},
                  @{Name = "LastLogonDate";Expression = {$_.lastlogondate}},
                  @{Name = "WhenUserWasCreated";Expression = {$_.whenCreated}},
                  @{Name = "accountexpiratondate";Expression = {$_.accountexpiratondate}},
                  @{Name = "PasswordLastSet";Expression = {([DateTime]::FromFileTime($_.pwdLastSet))}},
                  @{Name = "PasswordExpiryDate";Expression={([datetime]::fromfiletime($_."msDS-UserPasswordExpiryTimeComputed")).DateTime}},
                  @{Name = "Password Never";Expression = {$_.passwordneverexpire}},
                  @{Name = "HomeDriveLetter";Expression = {$_.HomeDrive}},
                  @{Name = "HomeFolder";Expression = {$_.HomeDirectory}},
                  @{Name = "scriptpath";Expression = {$_.scriptpath}},
                  @{Name = "HomePage";Expression = {$_.HomePage}},
                  @{Name = "Department";Expression = {$_.Department}},
                  @{Name = "EmployeeID";Expression = {$_.EmployeeID}},
                  @{Name = "Job Title";Expression = {$_.Title}},
                  @{Name = "EmployeeNumber";Expression = {$_.EmployeeNumber}},
                  @{Name = "Manager";Expression = {%{(Get-AdUser $_.Manager -Properties DisplayName).DisplayName}}}, 
                  @{Name = "Company";Expression = {$_.Company}},
                  @{Name = "Office";Expression = {$_.OfficeName}}
                  )
                  $Data | Sort LastLogondate -Descending |
                    Export-Csv -Path $path\ADUsers.csv -NoTypeInformation      
        
        
        #OUs
        
        Get-ADOrganizationalUnit -filter * | select Name,DistinguishedName,Description |
            Export-csv -path $path\ADOrganizationalUnits.csv -NoTypeInformation
        
        #contacts
        Get-ADobject  -LDAPfilter "objectClass=contact" -Properties mail,Description  | Select-Object name,mail,Description |
            Export-csv -path $path\ADContacts.csv -NoTypeInformation
        
        # AD Groups
        Get-ADgroup -Filter * -Properties members,whencreated,description,groupscope | select name,samaccountname,groupscope, @{n=’Members’; e= { ( $_.members | % { (Get-ADObject $_).Name }) -join “,” }},whencreated,description | Sort-Object -Property Name |
            Export-csv -path $path\ADGroups.csv -NoTypeInformation

        # Merge CSV files into XLSX
        <#Import the CSVs
        ## GC = Get-Content
        $csv1 = @(gc "$path\ADComputers.csv")
        $csv2 = @(gc "$path\ADUsers.csv")
        $csv3 = @(gc "$path\ADOrganizationalUnits.csv")
        $csv4 = @(gc "$path\ADContacts.csv")
        $csv5 = @(gc "$path\ADGroups.csv")
        # Create an Empty Array
        $csv6 = @()
            for ($i=0; $i -lt $csv1.Count; $i++) {
            $csv4 += $csv1[$i] + ', ' + $csv2[$i] + ', ' + $csv3[$i]
            }
        # Output to file
        $csv6 | Out-File "$path\AD-Merged.csv" -encoding default

        # Delete the originals
        #Remove-Item $csv1,$csv2,$csv3,$csv4,$csv5
#>

        #Create XLSX from CSV
        $csvs = Get-ChildItem $path\* -Include *.csv
        $y=$csvs.Count
            Write-Host "Detected the following CSV files: ($y)"
foreach ($csv in $csvs)
{
            Write-Host " "$csv.Name
}
        $outputfilename = $(get-date -f yyyyMMdd) + "_" + $env:USERNAME + "_combined-data.xlsx" #creates file name with date/username
            Write-Host Creating: $outputfilename
        $excelapp = new-object -comobject Excel.Application
        $excelapp.sheetsInNewWorkbook = $csvs.Count
        $xlsx = $excelapp.Workbooks.Add()
        $sheet=1

foreach ($csv in $csvs)
{
        $row=1
        $column=1
        $worksheet = $xlsx.Worksheets.Item($sheet)
        $worksheet.Name = $csv.Name
        $file = (Get-Content $csv)
foreach($line in $file)
{
        $linecontents=$line -split ',(?!\s*\w+")'
foreach($cell in $linecontents)
{
        $worksheet.Cells.Item($row,$column) = $cell
        $column++
}
        $column=1
        $row++
}
        $sheet++
}
        $output = $path + "\" + $outputfilename
        $xlsx.SaveAs($output)
        $excelapp.quit()