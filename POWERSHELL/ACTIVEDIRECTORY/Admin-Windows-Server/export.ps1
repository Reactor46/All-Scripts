$Date = Get-Date -UFormat "%Y_%m_%d_%H_%M"

$OutFile = "C:\Backup\Backup_$Date.csv"


if (Test-Path $OutFile){
    Del $OutFile
}


if (!(Test-Path -Path "C:\Backup")){
    New-Item -ItemType Directory -Path C:\Backup

}




$InputDN = Read-Host -Prompt "Write the DistinguishedName of the Organisation Unit"

Import-Module ActiveDirectory
set-location ad:

(Get-Acl $InputDN).access | ft identityreference, accesscontroltype, isinherited -autosize



$Childs = Get-ChildItem $InputDN -recurse


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
