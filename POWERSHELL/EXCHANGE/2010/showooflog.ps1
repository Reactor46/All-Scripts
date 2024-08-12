$MailboxName = "user@domain"
$rptCollection = @()

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$service.TraceEnabled = $false


$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
# $service.Credentials = New-Object System.Net.NetworkCredential("user","password")
$service.AutodiscoverUrl($MailboxName ,{$true})

$nonipmRoot = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
$ivItemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(2)
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Microsoft.OOF.Log")
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
$fiItems = $service.finditems($nonipmRoot,$sfSearchFilter,$ivItemView)
foreach($itItem in $fiItems.Items){

	$itItem.load($psPropertySet)
	$itItem.Body.Text.Split("`r") | ForEach-Object{
		$line = $_
		if($line.indexof("Mailbox:")-gt 0){
			$elen = $line.indexof("Mailbox:")
			if($elen -gt 0){
				$datetime = $line.SubString(1,$elen-3)
			}
		}
		if($line.indexof("Mailbox:")-gt 0){
			$slen = $line.indexof("Mailbox:")
			$elen = $line.indexof("OofState:",$slen)
			if($slen -gt 0){
				$slen += 10 
				$mailbox = $line.SubString($slen,($elen-$slen)-3)
			}
		}
		if($line.indexof("OofState:")-gt 0){
			$slen = $line.indexof("OofState:")
			$elen = $line.indexof("ExternalAudience:",$slen)
			if($slen -gt 0){
				$slen += 10
				$OofState = $line.SubString($slen,($elen-$slen)-2)
			}
		}
		if($line.indexof("ExternalAudience:")-gt 0){
			$slen = $line.indexof("ExternalAudience:")
			$elen = $line.indexof("InternalReply:",$slen)
			if($slen -gt 0){
				$slen += 18
				$ExternalAudience = $line.SubString($slen,($elen-$slen)-2)
			}
		}
		if($line.indexof("InternalReply:")-gt 0){
			$slen = $line.indexof("InternalReply:")
			$elen = $line.indexof("ExternalReply:",$slen)
			if($slen -gt 0){
				$slen += 15
				$InternalReply = $line.SubString($slen,($elen-$slen)-2)
			}
		}
		if($line.indexof("ExternalReply:")-gt 0){
			$slen = $line.indexof("ExternalReply:")
			$elen = $line.indexof("SetByLegacyClient:",$slen)
			if($slen -gt 0){
				$slen += 15
				$ExternalReply = $line.SubString($slen,($elen-$slen)-2)
			}
		}
		if($line.indexof("SetByLegacyClient:")-gt 0){
			$slen = $line.indexof("SetByLegacyClient:")
			$elen = $line.indexof("comment:")
			if($slen -gt 0){
				$slen += 20
				$SetByLegacyClient = $line.SubString($slen,($elen-$slen)-3)
			}
		}
		if($line.indexof("comment:")-gt 0){
			$slen = $line.indexof("comment")
			$elen = $line.indexof("'",$slen)
			if($slen -gt 0){
				$slen += 9
				$comment = $line.SubString($slen,($elen-$slen))
			}
		}
		$mbcomb = "" | select datetime,mailbox,OofState,ExternalAudience,InternalReply,ExternalReply,SetByLegacyClient,comment		
		$mbcomb.datetime = $datetime
		$mbcomb.mailbox = $mailbox
		$mbcomb.OofState = $OofState
		$mbcomb.ExternalAudience = $ExternalAudience
		$mbcomb.InternalReply = $InternalReply
		$mbcomb.ExternalReply = $ExternalReply
		$mbcomb.SetByLegacyClient = $SetByLegacyClient
		$mbcomb.comment = $comment
		$rptCollection += $mbcomb
	}
}
$rptCollection 