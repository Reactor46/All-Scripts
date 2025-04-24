###########################################################################
# Query list of IIS servers for Binding Information
###########################################################################
$serverlist = get-content "$($env:USERPROFILE)\Desktop\IISServerList.txt"
$csvexportfile = "$($env:USERPROFILE)\IISBindingsExport.csv"

$FullIISSiteList = @()

foreach ($server in $serverlist) {
   # All sites
   $IISSites = Invoke-Command -ComputerName $server -ScriptBlock { Get-WebBinding }
   foreach ($site in $IISSites) {
        # Split the site path into usable chunks.
        $sitePath = $site.ItemXPath -split "'"
        write-output $sitePath[1]

        # Create new object for injection into Array.
        $obj = New-Object PSObject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name ServerName -Value  $server
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Protocol -Value $site.protocol
        Add-Member -InputObject $obj -MemberType NoteProperty -Name PortBinding -Value $site.bindingInformation
        Add-Member -InputObject $obj -MemberType NoteProperty -Name SSL -Value $site.sslFlags
        Add-Member -InputObject $obj -MemberType NoteProperty -Name Path -Value $sitePath[1]
        Add-Member -InputObject $obj -MemberType NoteProperty -Name CertificateHash -Value $site.certificateHash
        $FullIISSiteList += $obj
   }
}

$FullIISSiteList | Export-Csv -Path $csvexportfile -NoTypeInformation