<<<<<<< HEAD
﻿#Fonction de recuperation de toutes les OU du domaine
Function Get-ADSIOU
{
    # Renvoie toutes les OU du domaine courrant
    $objDomain = [ADSI]''
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($objDomain)
    $objSearcher.Filter = '(objectCategory=organizationalUnit)'

    $OU = $objSearcher.FindAll() | Select-object -ExpandProperty Path
    
    $OU
}


#Fonction de recuperation de tous les comptes du domaine
Function Get-ADUsers
{
    echo "**** Utilisateurs du domaine*****"

    $ldapQuery = "(&(objectCategory=user))"
    $de = new-object system.directoryservices.directoryentry
    $ads = new-object system.directoryservices.directorysearcher -argumentlist $de,$ldapQuery
    $complist = $ads.findall()
    foreach ($i in $complist) 
    {
        write-host $i.Path
        
    }
}

Function Get-ADGroups
{
    echo "**** Groupes du domaine*****"

    $Groupes = 'Administrateur','Admin','Administrateurs','Admins'
    $ldapQuery = "(&(objectCategory=group))"
    $de = new-object system.directoryservices.directoryentry
    $ads = new-object system.directoryservices.directorysearcher -argumentlist $de,$ldapQuery
    $complist = $ads.findall()
    foreach ($i in $complist) 
    {
        write-host $i.Path
        if([boolean]( $i.Path.memberof  -match ($Groupes -join '|')) )
        {
            echo "Ce groupe fais partie des Administrateurs"
        }
        
    }
}

Function Get-ADComputers
{
    echo "**** Ordinateurs du domaine*****"
    $ldapQuery = "(&(objectCategory=computer))"
    $de = new-object system.directoryservices.directoryentry
    $ads = new-object system.directoryservices.directorysearcher -argumentlist $de,$ldapQuery
    $complist = $ads.findall()
    foreach ($i in $complist) 
    {
        write-host $i.Path
        
    }
}

echo "Welcome in this function tester"
$saisie = 2
while($saisie -lt 6 -and $saisie -gt 0)
{
    echo "******MENU******"
    echo "1- Get-ADUsers"
    echo "2- Get-ADComputers"
    echo "3- Get-ADGroups"
    echo "4- Get-ADSIOU"
    echo "5- Leave the script"
    $saisie=Read-Host ">>> "

    if($saisie -eq 1)
    {
        Get-ADUsers
    }
    elseif($saisie -eq 2)
    {
        Get-ADComputers
    }
    elseif($saisie -eq 3)
    {
        Get-ADGroups
    }
    elseif($saisie -eq 4)
    {
        Get-ADSIOU
    }
    elseif($saisie -eq 5)
    {
        $saisie = 0
    }
}

$input = Read-Host -prompt "write to Ou name affected"
=======
﻿$input = Read-Host -prompt "write to Ou name affected"
>>>>>>> master

$allowance = Read-Host -Prompt "Deny / Allow "


$Ou = New-ADOrganizationalUnit -name $input -PassThru
$Acl = get-acl $Ou

## Note that bf967a86-0de6-11d0-a285-00aa003049e2 is the schemaIDGuid for the computer class.
# bf967aba-0de6-11d0-a285-00aa003049e2 is the schemaIDGuid for the user class.
#The following object specific ACE is to grant $Ou permission to create computer objects under $Ou.
 $computer_class_guid = new-object Guid bf967a86-0de6-11d0-a285-00aa003049e2
 $user_class_guid = new-object Guid  00299570-246d-11d0-a768-00aa006e0529

 $ace1 = new-object System.DirectoryServices.ActiveDirectoryAccessRule $Ou,"CreateChild","Allow",$computer_class_guid

 $acl.AddAccessRule(($ace1)
Set-acl -Path -AclObject $acl $Ou