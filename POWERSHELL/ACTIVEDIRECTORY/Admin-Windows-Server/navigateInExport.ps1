#Variable Globale
$GLOBAL:ADPATH = $null

Function Navigate
{
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry


    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher

    $objSearcher.SearchRoot = $objDomain

    $objSearcher.Filter = ("(objectCategory=organizationalUnit)")


    $colProplist = "name"

    foreach ($i in $colPropList)
        {
            $objSearcher.PropertiesToLoad.Add($i)
        }

    write-host "Here are the differents Organizational Units of the Domain : $Global:DOMAIN"
    write-host "------------------------------------"

    $colResults = $objSearcher.FindAll()
    $cpt = 0
    foreach ($objResult in $colResults)
        {
            $objComputer = $objResult.Properties;

            $name = $objComputer.name
            write-host "$cpt -> $name"
            $cpt = $cpt + 1
        }
    write-host "------------------------------------"
    $nb_ou = Read-Host "Select your number "
    $cpt = 0
    foreach ($objResult in $colResults)
        {
            $objComputer = $objResult.Properties;

            $name = $objComputer.name
            if($cpt -eq $nb_ou)
            {
                $input = $name
                $LDAPway = $objResult.Path
            }
            $cpt = $cpt + 1
        }

    $stop = 0
    while($stop -eq 0)
    {
        write-host "Here are the childs of the Organizational Unit you selected : $input"
        write-host "------------------------------------"
        Write-Host "LDAP =$LDAPway"
        $nb_char = $LDAPway.Length

        $OUdef = $LDAPway.Substring(7,$nb_char-7)
        $ADPath = "AD:\$OUdef"
        $OU_childs = Get-ChildItem -Path $ADPath

        $cpt = 0
        foreach($ou_child in $OU_childs)
        {
            $name = $ou_child.name
            write-host "[$cpt] -> $name"
            $cpt = $cpt + 1
        }
        write-host "-1 -> QUIT"
        write-host "------------------------------------"
        $nb_ou = Read-Host "Select your number "
        if($nb_ou -eq -1){$stop = 1}
        $cpt = 0
        foreach ($ou_child in $OU_childs)
            {
                $ou_chil = $ou_child.Properties;

                if($cpt -eq $nb_ou)
                {
                    $tmp = $ou_child.name
                    $input = "$input >> $tmp"
                    $LDAPway = "LDAP://$ou_child"
                }
                $cpt = $cpt + 1
            }
    }

    $Global:ADPATH = $ADPath.ToString()
}

$Date = Get-Date -UFormat "%Y_%m_%d_%H_%M"

$OutFile = "C:\Backup\Backup_$Date.csv"


if (Test-Path $OutFile){
    Del $OutFile
}


if (!(Test-Path -Path "C:\Backup")){
    New-Item -ItemType Directory -Path C:\Backup

}




Navigate

$ADPath = $GLOBAL:ADPATH

$nb_char = $GLOBAL:ADPATH.Length
$tmp = $GLOBAL:ADPATH.Substring(4,$nb_char-4)
$InputDN = 'LDAP://'+ $tmp


Import-Module ActiveDirectory
set-location ad:

(Get-Acl -Path $ADPath).access | ft identityreference, accesscontroltype, isinherited -autosize



$Childs = Get-ChildItem -Path $ADPath -recurse


foreach($Child in $Childs){


    Write-Host $Child.distinguishedName

    $Header = $Child.distinguishedName

    Add-Content -Value $Header -Path $OutFile


    $Header = "IdentityReference,AccessControlType,IsInherited"
    Add-Content -Value $Header -Path $OutFile





    (Get-Acl $Child.DistinguishedName).access | ft identityreference, accesscontroltype, isinherited -autosize

     $ACLs = Get-Acl $Child.DistinguishedName | ForEach-Object {$_.access}




    Foreach ($ACL in $ACLs){
	    $OutInfo = $ACL.identityreference


       if ($ACL.AccessControlType -eq "Allow"){
            $OutInfo = "$OutInfo, Allow"

        } else {
            $OutInfo = "$OutInfo, Deny"
        }


        if ($ACL.IsInherited -eq "True"){
            $OutInfo = "$OutInfo, True"

        } else {
            $OutInfo = "$OutInfo, False"
        }



	    Add-Content -Value $OutInfo -Path $OutFile
	}


}
