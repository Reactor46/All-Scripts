#Allow the script execution
Set-ExecutionPolicy Unrestricted

#Set-Location AD

Function Get-LDAPway($chx)
{
    if($chx -eq 1)
    {
        $strCategory = "user"
    }
    elseif($chx -eq 2)
    {
        $strCategory = "computer"
    }
    elseif($chx -eq 3)
    {
        $strCategory = "group"
    }
    elseif($chx -eq 4)
    {
        $strCategory = "organizationalUnit"
    }

    if($chx -eq 4)
    {
        $objDomain = [ADSI]''
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher($objDomain)
        $objSearcher.Filter = '(objectCategory=$strCategory)'

        $OU = $objSearcher.FindAll() | Select-object -ExpandProperty Path

        $OU
    }
    else
    {
        $ldapQuery = "(&(objectCategory=$strCategory))"
        $de = new-object system.directoryservices.directoryentry
        $ads = new-object system.directoryservices.directorysearcher -argumentlist $de,$ldapQuery
        $complist = $ads.findall()
        foreach ($i in $complist) 
        {
            write-host $i.Path
            if($chx -ne 3)
            {
                echo ">>> Is member of :"
                ([ADSI]$i.Path).memberof
                echo "-------------------"
            }   
        }
    }
}

Function Get-ADObject($chx)
{
    if($chx -eq 1)
    {
        $strCategory = "user"
    }
    elseif($chx -eq 2)
    {
        $strCategory = "computer"
    }
    elseif($chx -eq 3)
    {
        $strCategory = "group"
    }
    elseif($chx -eq 4)
    {
        $strCategory = "organizationalUnit"
    }

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry


    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher

    $objSearcher.SearchRoot = $objDomain

    $objSearcher.Filter = ("(objectCategory=$strCategory)")


    $colProplist = "name","cn"

    foreach ($i in $colPropList)
        {
            $objSearcher.PropertiesToLoad.Add($i)
        }


    $colResults = $objSearcher.FindAll()


    foreach ($objResult in $colResults)
        {
            $objComputer = $objResult.Properties; 
            $objComputer.name
        }
}


echo "Welcome in this function test"
$saisie = 2
while($saisie -lt 7 -and $saisie -gt 0)
{
    echo "******MENU******"
    echo "1- Get-ADUsers"
    echo "2- Get-ADComputers"
    echo "3- Get-ADGroups"
    echo "4- Get-ADSIOU"
    echo "5- Function All In One"
    echo "6- Leave the script"
    $saisie=Read-Host ">>> "

    if($saisie -eq 1)
    {
        #Commande pour recuperer le DN du compte utilisé actuellement
        #([ADSI]"LDAP://$(whoami /fqdn)").memberof
        echo "******Nom de chaque objet de ce type****"
        Get-ADObject(1)
        echo "******Chemin LDAP de chaque objet de ce type******"
        Get-LDAPway(1)
    }
    elseif($saisie -eq 2)
    {
        echo "******Nom de chaque objet de ce type****"
        Get-ADObject(2)
        echo "******Chemin LDAP de chaque objet de ce type******"
        Get-LDAPway(2)
    }
    elseif($saisie -eq 3)
    {
        echo "******Nom de chaque objet de ce type****"
        Get-ADObject(3)
        echo "******Chemin LDAP de chaque objet de ce type******"
        Get-LDAPway(3)
    }
    elseif($saisie -eq 4)
    {
        echo "******Nom de chaque objet de ce type****"
        Get-ADObject(4)
        echo "******Chemin LDAP de chaque objet de ce type******"
        Get-LDAPway(4)
    }
    elseif($saisie -eq 5)
    {
        echo "----------------------------"
        echo "Get-ADObject . users . name"
        echo "----------------------------"
        Get-LDAPway(1)
        Get-ADObject(1)
        echo "----------------------------"
        echo "Get-ADObject . computers . name"
        echo "----------------------------"
        Get-LDAPway(2)
        Get-ADObject(2)
        echo "----------------------------"
        echo "Get-ADObject . groups . name"
        echo "----------------------------"
        Get-LDAPway(3)
        Get-ADObject(3)
        echo "----------------------------"
        echo "Get-ADObject . ou . name"
        echo "----------------------------"
        Get-LDAPway(4)
        Get-ADObject(4)
    }
    elseif($saisie -eq 6)
    {
        $saisie = 0
    }
}

#Get-ADUser -Identity 'CN=Petitjean Arnaud,CN=Users,DC=powershell-scripting,DC=com' -Properties Description

# Connexion à l'objet en spécifiant son DN - Distinguished Name
$user = [ADSI]'LDAP://CN=Gege GF. Firth,CN=Users,DC=esgi,DC=priv'

echo "$user"

# Modification de la propriété Description avec la méthode Put
#$user.Put('Description','Cet utilisateur est exceptionnel !')

# Application des changements avec la méthode SetInfo
#$user.SetInfo()

#Recuperation du SID d'un compte
#param ($account = $(throw "esgi.priv\Administrateur")) 

#if ($account -is [security.principal.ntaccount]) { 
#    $ntaccount = $account 

#} else {
#    $ntaccount = new-object security.principal.ntaccount $account 
#}

#$ntaccount.translate( [security.principal.securityidentifier] )