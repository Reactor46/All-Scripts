# Used to generate passwords when creating AD users.
function GeneratePassword() {
    $digits = 48..57
    $letters = 65..90 + 97..122
    
    $password = Get-Random -count 16 `
        -input ($digits + $letters) |
            % -begin { $aa = $null } `
            -process { $aa += [char]$_ } `
            -end { $aa }
    return "$password$(Get-Random -Maximum 9)"
}

# Gets domain and splits it in two...
$domain = Read-Host "Domain (dev.local)"
$prefix,$suffix = $domain.split('.', 2)

#Setting up SMTP for report...
$email = Read-Host "Office 365 Email Address"
$pass = Read-Host "Office 365 Password" -AsSecureString
$smtp = "smtp.office365.com"
$smtpcred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $email, $pass

$users = Import-CSV "users.csv"
$organizationalunits = Import-CSV "organizationalunits.csv"
$securitygroups = Import-CSV "securitygroups.csv"

### Adding OUs.

$report += "# Organizational Units`n"

ForEach($ou in $organizationalunits) {
    $ouName = $ou.Name
    
    if(![adsi]::Exists("LDAP://OU=$ouName,DC=$prefix,DC=$suffix")) {
        New-ADOrganizationalUnit -Name "$ouName" -Path "DC=$prefix,DC=$suffix"
        $report += "`n_$ouName"
    }
}

# Setting default OUs for users and computers

redircmp "OU=_Domain Computers,DC=$prefix,DC=$suffix"
redirusr "OU=_Domain Users,DC=$prefix,DC=$suffix"

### Adding Security Groups

$report += "`n`n# Security Groups`n"

ForEach($group in $securitygroups) {
    $grpName = $group.Name
    
    # Using a try/catch as without, it'll throw errors (though will still work). It's cleaner.
    try { Get-ADGroup "$grpName" }
    catch {
        try {
            New-ADGroup -Name $grpName -GroupScope Global `
            -Path "OU=_Security Groups,DC=$prefix,DC=$suffix"

            $report += "`n$grpName"
        }
        catch {
            $report += "`n$grpName... failed"
        }
    }
}

### Adding Users

$report += "`n`n# Users`n"

ForEach($user in $users) {
    # Creating information for user
    $name = $user.Name
    $givenName = $user.GivenName
    $surName = $user.SurName
    $samaccountname = $user.SamAccountName
    $userprincipalname = "$samaccountname@$domain"
    $ou = $user.OU
    $securitygroups = $user.SecurityGroups -Split ", "
    $password = GeneratePassword
    
    # Adding users
    New-ADUser -Name $name -GivenName $GivenName -Surname $SurName `
        -SamAccountName $samaccountname `
        -UserPrincipalName $userprincipalname -Enabled $true -PasswordNeverExpires $true `
        -AccountPassword (ConvertTo-SecureString -String $password -AsPlainText -Force) `
        -Path "OU=$ou,DC=$prefix,DC=$suffix"
        
    $report += "`nUsername: $samaccountname"
    $report += "`nPassword: $password`n"

    ForEach($sg in $securitygroups) {
        $report += "`n$samaccountname -> $sg"
        Add-ADGroupMember $sg $samaccountname
    }

    $report += "`n---`n"
}

Send-MailMessage -To $email -Subject "Domain Onboarding: $domain" -Body $report -From $email -SmtpServer $smtp -usessl -Credential $smtpcred
$report | Out-File "report.txt"