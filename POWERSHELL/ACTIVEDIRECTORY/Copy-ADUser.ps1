Param (
    
    [Parameter(Mandatory=$True)]
    $givenname,
    [Parameter(Mandatory=$True)]
    $surname,
    
    [Parameter(Mandatory=$True)]    
    $template,
    $password = "Newpass1",
    [switch]$enabled,
    $changepw = $true,
    $ou,
    [switch]$useTemplateOU
)
$name = "$givenname $surname"
$samaccountname = "$($givenname[0])$surname"
$password_ss = ConvertTo-SecureString -String $password -AsPlainText -Force
$template_obj = Get-ADUser -Identity $template
If ($useTemplateOU) {
    $ou = $template_obj.DistinguishedName -replace '^cn=.+?(?<!\\),'
}
$params = @{
    "Instance"=$template_obj
    "Name"=$name
    "DisplayName"=$name
    "GivenName"=$givenname
    "SurName"=$surname
    "AccountPassword"=$password_ss
    "Enabled"=$true
    "ChangePasswordAtLogon"=$changepw
}
If ($ou) {
    $params.Add("Path",$ou)
}
New-ADUser @params