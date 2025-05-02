
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$snStickyNote = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
$snStickyNote.Subject = "Racing"
$snStickyNote.ItemClass = "IPM.StickyNote"
$snStickyNote.Body = "First Line `nNext Line"

$noteGuid = new-object Guid("0006200E-0000-0000-C000-000000000046")
$snColour = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($noteGuid,35584, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$snHeight = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($noteGuid,35586, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$snWidth = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($noteGuid,35587, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$snLeft = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($noteGuid,35588, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$snTop = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($noteGuid,35589, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)

$snStickyNote.ExtendedProperties.Add($snColour,3)
$snStickyNote.ExtendedProperties.Add($snHeight,200)
$snStickyNote.ExtendedProperties.Add($snWidth,166)
$snStickyNote.ExtendedProperties.Add($snLeft,80)
$snStickyNote.ExtendedProperties.Add($snTop,80)

$snStickyNote.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Notes)
