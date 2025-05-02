[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

$snServerName = $args[0]
$username = "userName"
$password = "password"
$domain = "domain"


$Global:rptCollection = @()

## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624



function QueryMailbox($mbURI){
$datetimetoquery = get-date
write-host $mbURI
$wdWebDAVQuery = "<?xml version=`"1.0`"?><D:searchrequest xmlns:D = `"DAV:`" ><D:sql>
SELECT `"DAV:childcount`" As Count,`"http://schemas.microsoft.com/mapi/proptag/x0e080003`" As Size FROM scope('shallow
traversal of `"" + $mbURI + "`"')
Where `"DAV:ishidden`" = False AND `"DAV:isfolder`" = True AND `"DAV:displayname`" = 'Calendar'</D:sql></D:searchrequest>"

$WDRequest = [System.Net.WebRequest]::Create($mbURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "SEARCH"
$Credential = New-Object System.Net.NetworkCredential($user, $password, $domain)
$CCache = New-Object System.Net.CredentialCache
$CCache.Add($mbURI, "Basic", $Credential)
$WDRequest.Credentials = $CCache
$bytes = [System.Text.Encoding]::UTF8.GetBytes($wdWebDAVQuery)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$SizeNodes = @($ResponseXmlDoc.getElementsByTagName("Size"))
$ItemCountNodes = @($ResponseXmlDoc.getElementsByTagName("Count"))
$Itmcnt = "" | select DisplayName,SMTPAddress,ItemCount,Size
$Itmcnt.DisplayName = $displayName
$Itmcnt.SMTPAddress = $SmtpAddress
$Itmcnt.ItemCount = $ItemCountNodes[0].'#text'
$Itmcnt.Size = [System.Math]::Round(($SizeNodes[0].'#text' /1024),2)
$Itmcnt
$Global:rptCollection += $Itmcnt
}

function GetUsers(){
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$dfRoot = [ADSI]$dfDefaultRootPath
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(&(objectCategory=msExchExchangeServer)(cn=' + $snServerName  + '))'
$searcher.PropertiesToLoad.Add("cn")
$searcher.PropertiesToLoad.Add("gatewayProxy")
$searcher.PropertiesToLoad.Add("legacyExchangeDN")
$searcher1 = $searcher.FindAll()
foreach ($server in $searcher1){ 
	$snServerEntry = New-Object System.DirectoryServices.directoryentry 
        $snServerEntry = $server.GetDirectoryEntry() 
	$snServerName = $snServerEntry.cn
	$snExchangeDN = $snServerEntry.legacyExchangeDN
}
$searcher.Filter = '(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy))'
$searcher1 = $searcher.FindAll()
foreach ($recppolicies in $searcher1){ 
     	$gwaddrrs = New-Object System.DirectoryServices.directoryentry 
        $gwaddrrs = $recppolicies.GetDirectoryEntry() 
	foreach ($address in $gwaddrrs.gatewayProxy){
		if($address.Substring(0,5) -ceq "SMTP:"){$dfAddress = $address.Replace("SMTP:@","")}
	}	
	
}
$arMbRoot = "https://" + $snServerName + "/exadmin/admin/" + $dfAddress + "/mbx/"
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)" `
+ "(objectClass=user)(msExchHomeServerName=" + $snExchangeDN + ")) )))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$searcher2 = $dfsearcher.FindAll()
foreach ($uaUsers in $searcher2){ 
     	$uaUserAccount = New-Object System.DirectoryServices.directoryentry 
        $uaUserAccount = $uaUsers.GetDirectoryEntry() 
	foreach ($address in $uaUserAccount.proxyaddresses){
		if($address.Substring(0,5) -ceq "SMTP:"){$uaAddress = $address.Replace("SMTP:","")}
	}
	$SmtpAddress = $uaAddress
	$displayName = $uaUserAccount.DisplayName[0]
	QueryMailbox($arMbRoot + $uaAddress + "")
}
}

GetUsers
$Global:rptCollection | Export-Csv -NoTypeInformation c:\mailbox.csv