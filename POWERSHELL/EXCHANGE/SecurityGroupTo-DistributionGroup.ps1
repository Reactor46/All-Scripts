New-ADGroup -Name "TAB_INTERACTOR_DL" -SamAccountName "TAB_INTERACTOR_DL" -GroupCategory Distribution -GroupScope Universal -DisplayName "TAB_INTERACTOR_DL" -Path "OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com" -Description "Tableau Interactor distribution group"
New-ADGroup -Name "TAB_PUBLISHER_DL" -SamAccountName "TAB_PUBLISHER_DL" -GroupCategory Distribution -GroupScope Universal -DisplayName "TAB_PUBLISHER_DL" -Path "OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com" -Description "Tableau Publisher distribution group"
New-ADGroup -Name "TAB_SERVER_ADMIN_DL" -SamAccountName "TAB_SERVER_ADMIN_DL" -GroupCategory Distribution -GroupScope Universal -DisplayName "TAB_SERVER_ADMIN_DL" -Path "OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com" -Description "Tableau Server Admin distribution group"

@"
SrcGrp, DstGrp
"CN=TAB_INTERACTOR,OU=Risk Security Groups,OU=Security_Groups,OU=Las_Vegas,DC=contoso,DC=com","CN=TAB_INTERACTOR_DL,OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com"
"CN=TAB_PUBLISHER,OU=Risk Security Groups,OU=Security_Groups,OU=Las_Vegas,DC=contoso,DC=com","CN=TAB_PUBLISHER_DL,OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com"
"CN=TAB_SERVER_ADMIN,OU=Risk Security Groups,OU=Security_Groups,OU=Las_Vegas,DC=contoso,DC=com","CN=TAB_SERVER_ADMIN_DL,OU=Distribution_Lists,OU=Las_Vegas,DC=contoso,DC=com"
"@ | convertFrom-CSV |
  % { Add-ADGroupMember $_.DstGrp -Members (Get-ADGroupMember $_.SrcGrp) }