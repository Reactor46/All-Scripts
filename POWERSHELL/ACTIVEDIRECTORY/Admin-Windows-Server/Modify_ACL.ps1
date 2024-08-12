Function Add_ACL
{
    write-host "These are the differents Organizational Units on this domain"
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry


    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher

    $objSearcher.SearchRoot = $objDomain

    $objSearcher.Filter = ("(objectCategory=organizationalUnit)")


    $colProplist = "name"

    foreach ($i in $colPropList)
        {
            $objSearcher.PropertiesToLoad.Add($i)
        }


    $colResults = $objSearcher.FindAll()
    $cpt = 0
    foreach ($objResult in $colResults)
        {
            $objComputer = $objResult.Properties;

            $name = $objComputer.name
            write-host "$cpt : $name"
            $cpt = $cpt + 1
        }
    write-host "------------------------------------"
    write-host "------------------------------------"
    write-host "Please write the number of the Organizational Unit you want"
    $nb_ou = Read-Host ">>> "
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

    write-host "OU choisis : $input - $LDAPway"


    $nb_char = $LDAPway.Length

    $OUdef = $LDAPway.Substring(7,$nb_char-7)
    $ADPath = "AD:\$OUdef"

    write-host $ADPath



    Try{
        $ou = Get-ADOrganizationalUnit -LDAPFilter $LDAPway -Credential $GLOBAL:CRED
    }
    Catch{
        Write-Host "Error: OU was not found: " $_.Exception.Message -BackgroundColor Black -ForegroundColor Red
        $inpustop = Read-Host "..."
        Break
    }

    Try{
        $objACL = (Get-Acl -Path $ADPath).Access | ? ActiveDirectoryRights 
        }
    Catch{
        Write-Host "Error: ACL was not found in OU: " $_.Exception.Message -BackgroundColor Black -ForegroundColor Red
        $inpustop = Read-Host " ..."
        Break
    }

    $objDomain = [ADSI]$LDAPway
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.Filter = ("(objectCategory=group)")

    $colProplist = "name"

    foreach ($i in $colPropList)
        {
            $objSearcher.PropertiesToLoad.Add($i)
        }

    $groups = $objSearcher.FindAll()

    $cpt = 0
    write-host "Here are the differents groups member of this OU"
    foreach ($group in $groups)
        {
            $objgroup = $group.Properties;
            $name = $objgroup.name
            write-host "$cpt - $name " 

            $cpt = $cpt + 1
        }

    write-host "Do you want to Allow or Deny the next selection ?"
    $rightType = Read-Host "1- Allow | 2- Deny  >>> "


$cptr = 0
$cptwr = 0
$cptrwr = 0
$cptex = 0
$cptrex = 0
$cptdel = 0
$cptmodif = 0
$cptcrea = 0
$cptcreadir = 0
$cptfullc = 0
$cptAll = 0
$cptlistch = 0
$cptlistob = 0
$cptsynch = 0
$cptwrprop = 0

$cptchoose = 0


$right = 1
while($right -gt 0 -and $right -lt 16)
{
    if($rightType -eq 1)
    {
        $rightype =[System.Security.AccessControl.AccessControlType]::Allow
    }
    elseif($rightType -eq 2)
    {
        $rightype = [System.Security.AccessControl.AccessControlType]::Deny
    }

    write-host "Here are different rights you can choose to apply on these groups" 
    if($cptr -eq 0){write-host "1- GenericRead"}
    if($cptwr -eq 0){write-host "2- GenericWrite"}
    if($cptrwr -eq 0){write-host "3- ReadControl"}
    if($cptex -eq 0){write-host "4- GenericExecute"}
    if($cptrex -eq 0){write-host "5- ReadProperty"}
    if($cptdel -eq 0){write-host "6- Delete"}
    if($cptmodif -eq 0){write-host "7- DeleteChild"}
    if($cptcrea -eq 0){write-host "8- DeleteTree"}
    if($cptcreadir -eq 0){write-host "9- CreateChild"}
    if($cptfullc -eq 0){write-host "10- WriteDacl"}
    if($cptAll -eq 0){write-host "11- GenericAll"}
    if($cptlistch -eq 0){write-host "12- ListChildren"}
    if($cptlistob = 0 -eq 0){write-host "13- ListObject"}
    if($cptsynch -eq 0){write-host "14- Synchronize"}
    if($cptwrprop -eq 0){write-host "15- WriteProperty"}
    write-host "16- Stop"
    $right = $null
    $right=Read-Host ">>> "
  
    if($right -ne 16 -and $cptchoose -gt 0)
    {
        $colRights = "$colRights,"
    }

    if($right -eq 1 -and $cptr -eq 0)
    {
        $colRights = $colRights + "GenericRead"
        $cptr = $cptr + 1 
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 2 -and $cptwr -eq 0)
    {        
        $colRights = $colRights + "GenericWrite" 
        $cptwr = $cptwr + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 3 -and $cptrwr -eq 0)
    {        
        $colRights = $colRights + "ReadControl" 
        $cptrwr = $cptrwr + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 4 -and $cptex -eq 0)
    {        
        $colRights = $colRights + "GenericExecute" 
        $cptex = $cptex + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 5 -and $cptrex -eq 0)
    {        
        $colRights = $colRights + "ReadProperty"
        $cptrex = $cptrex + 1
        $cptchoose = $cptchoose + 1 
    }
    elseif($right -eq 6 -and $cptdel -eq 0)
    {        
        $colRights = $colRights + "Delete"
        $cptdel = $cptdel + 1 
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 7 -and $cptmodif -eq 0)
    {
        $colRights = $colRights + "DeleteChild"
        $cptmodif = $cptmodif + 1 
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 8 -and $cptcrea -eq 0)
    {
        $colRights = $colRights + "DeleteTree"
        $cptcrea = $cptcrea + 1 
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 9 -and $cptcreadir -eq 0)
    {
        $colRights = $colRights + "CreateChild" 
        $cptcreadir = $cptcreadir + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 10 -and $cptfullc -eq 0)
    {
        $colRights = $colRights + "WriteDacl" 
        $cptfullc = $cptfullc + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 11 -and $cptAll -eq 0)
    {
        $colRights = $colRights + "GenericAll" 
        $cptAll = $cptAll + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 12 -and $cptlistch -eq 0)
    {
        $colRights = $colRights + "ListChildren" 
        $cptlistch = $cptlistch + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 13 -and $cptlistob -eq 0)
    {
        $colRights = $colRights + "ListObject" 
        $cptlistob = $cptlistob + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 14 -and $cptsynch -eq 0)
    {
        $colRights = $colRights + "Synchronize" 
        $cptsynch = $cptsynch + 1
        $cptchoose = $cptchoose + 1
    }
    elseif($right -eq 15 -and $cptwrprop -eq 0)
    {
        $colRights = $colRights + "WriteProperty" 
        $cptwrprop = $cptwrprop + 1
        $cptchoose = $cptchoose + 1
    }

    $ObjectGuid = [GUID]("bf967a86-0de6-11d0-a285-00aa003049e2")

    foreach($group in $groups)
    {
        $objgroup = $group.Properties;
        $name = $objgroup.name
        $Group = "$Global:DOMAIN\$name"
        write-host $Group
        $aceID = New-Object System.Security.Principal.NTAccount($Group)
        $ACE = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule($aceID,$colRights,$rightype,$ObjectGUID,1)


        #$ACE TOUJOURS NULL ??!!!
        $ADAcl.AddAccessRule($ACE)
    }


    Try
    {
        Set-ACL -Path $ADPath -ACLObject $ADAcl -Passthru 
    }
    Catch
    {
        Write-Host "Error: Set-ACL didn't work: " $_.Exception.Message -BackgroundColor Black -ForegroundColor Red
        $inpustop = Read-Host "..."
        Break
    }
    $inpustop = Read-Host "..."
}