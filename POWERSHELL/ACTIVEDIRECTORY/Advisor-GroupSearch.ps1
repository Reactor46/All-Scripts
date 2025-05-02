$Advisors = GC .\Advisors.txt

<#
ForEach($advisor in $Advisors){
    List-Groups $advisor | Export-CSV -Path .\AdvisorGroupExports\$advisor.csv -NoTypeInformation
    }
#>
ForEach($advisor in $Advisors){
    Get-ADPrincipalGroupMembership -Identity $advisor | Select -ExpandProperty Name | Out-File -FilePath .\AdvisorGroupExports\$advisor.txt -Encoding utf8
    }