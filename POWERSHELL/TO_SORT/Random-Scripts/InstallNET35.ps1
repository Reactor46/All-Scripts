Import-DscResource –ModuleName 'PSDesiredStateConfiguration'
$ConfigurationData = @{
    AllNodes = @(
        @{
            NodeName = '$env:COMPUTERNAME'
            NetFramework35 = $False
            Administrators  = @('KELSEY-SEYBOLD\WebAdmins')
        }
    )
}

Configuration SiteCore
{

    Node $AllNodes.NodeName

    {
        If($Node.NetFramework35 -eq $True)
        {
            WindowsFeature 'NetFramework35'

            {
                Name   = 'NET-Framework-Core'
                Source = '\\kscfs\it-misc$\IT_ESS_Team\KSC-ESS-DSC\SXS-2019'
                Ensure = 'Present'
            }
        }

	  Group 'Groups(XML): Administrators (built-in)'
          {
               MembersToInclude  = $node.Administrators
               GroupName = 'Administrators'

          }
    }
}

SiteCore -ConfigurationData $ConfigurationData

Start-DscConfiguration -Path C:\Temp\Sitecore\ -wait -verbose
Remove-Item .\SiteCore\ -Force:$True -Confirm:$False -Recurse