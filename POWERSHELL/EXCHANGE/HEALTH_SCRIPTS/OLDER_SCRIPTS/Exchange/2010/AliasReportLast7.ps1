## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

$Mailbox = get-mailbox $MailboxName
$mbAddress = @{}
foreach($emEmail in $Mailbox.EmailAddresses){
	if($emEmail.Tolower().Contains("smtp:")){
		$mbAddress.Add($emEmail.SubString(5).Tolower(),0)
	}
}

## Load Managed API dll  

###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL)
    {
    Import-Module $EWSDLL
    }
else
    {
    "$(get-date -format yyyyMMddHHmmss):"
    "This script requires the EWS Managed API 1.2 or later."
    "Please download and install the current version of the EWS Managed API from"
    "http://go.microsoft.com/fwlink/?LinkId=255472"
    ""
    "Exiting Script."
    exit
    }
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
$service.Credentials = $creds      
  
#Credentials Option 2  
#service.UseDefaultCredentials = $true  
  
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
  
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
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$PR_TRANSPORT_MESSAGE_HEADERS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x007D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
$psPropset.Add($PR_TRANSPORT_MESSAGE_HEADERS)
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
$ivItemView.PropertySet = $psPropset
$rptCollection = @()
$rptcntCount = @{}
#Find Items that are unread
$sfItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, (Get-Date).AddDays(-7)); 
$fiItems = $null    
do{    
    $fiItems = $service.FindItems($Inbox.Id,$sfItemSearchFilter,$ivItemView)    
    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){    
		$Headers = $null;
		$Address = "";
		$Method = ""
		if($Item.TryGetProperty($PR_TRANSPORT_MESSAGE_HEADERS,[ref]$Headers)){
			$slen = $Headers.ToLower().IndexOf("`nto: ")
			if($slen -gt 0){
				$elen = $Headers.IndexOf("`r`n",$slen)				
				$ToAddressParsed = $Headers.SubString(($slen+4),$elen-($slen+4))
				while($Headers.Substring($elen+2,1) -eq [char]9){
					$slen = $elen+2
					$elen = $Headers.IndexOf("`r`n",$slen)	
					$ToAddressParsed += $Headers.SubString(($slen),$elen-($slen))
				}
				$emailRegex = new-object System.Text.RegularExpressions.Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase);
            	$emailMatches = $emailRegex.Matches($ToAddressParsed);
				foreach($Match in $emailMatches){
					if($mbAddress.Contains($Match.Value.Tolower())){
						$Address = $Match
						$Method = "To"
						if($rptcntCount.ContainsKey($Match.Value)){
							$rptcntCount[$Match.Value].To +=1
						}
						else{
							$rptObj = "" | Select Address,To,CC,BCC
							$rptObj.Address = $Match.Value
							$rptObj.To = 1
							$rptObj.CC = 0
							$rptObj.BCC = 0
							$rptcntCount.Add($Match.Value,$rptObj)					
						}
					}
					else{
						if($Address -eq ""){
							$Address = $Match
							$Method = "Other"
						}
					}
				}
			}
			$slen = $Headers.ToLower().IndexOf("`ncc: ")
			if($slen -gt 0){
				$elen = $Headers.IndexOf("`r`n",$slen)
				$ccAddressParsed = $Headers.SubString(($slen+4),$elen-($slen+4))
				while($Headers.Substring($elen+2,1) -eq [char]9){
					$slen = $elen+2
					$elen = $Headers.IndexOf("`r`n",$slen)	
					$ccAddressParsed += $Headers.SubString(($slen),$elen-($slen))
				}
				$emailRegex = new-object System.Text.RegularExpressions.Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase);
            	$emailMatches = $emailRegex.Matches($ccAddressParsed);
				foreach($Match in $emailMatches){
					if($mbAddress.Contains($Match.Value.Tolower())){
						$Address = $Match
						$Method = "CC"
						if($rptcntCount.ContainsKey($Match.Value)){
							$rptcntCount[$Match.Value].CC +=1
						}
						else{
							$rptObj = "" | Select Address,To,CC,BCC
							$rptObj.Address = $Match.Value
							$rptObj.To = 0
							$rptObj.CC = 1
							$rptObj.BCC = 0
							$rptcntCount.Add($Match.Value,$rptObj)					
						}
					}
				}
			}
			$slen = $Headers.ToLower().IndexOf("`nbcc: ")
			if($slen -gt 0){
				$elen = $Headers.IndexOf("`r`n",$slen)
				$bccAddressParsed = $Headers.SubString(($slen+4),$elen-($slen+4))
				while($Headers.Substring($elen+2,1) -eq [char]9){
					$slen = $elen+2
					$elen = $Headers.IndexOf("`r`n",$slen)	
					$bccAddressParsed += $Headers.SubString(($slen),$elen-($slen))
				}
				$emailRegex = new-object System.Text.RegularExpressions.Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase);
            	$emailMatches = $emailRegex.Matches($bccAddressParsed);
				foreach($Match in $emailMatches){
					if($mbAddress.Contains($Match.Value.Tolower())){
						$Address = $Match
						$Method = "BCC"
						if($rptcntCount.ContainsKey($Match.Value)){
							$rptcntCount[$Match.Value].BCC +=1
						}
						else{
							$rptObj = "" | Select Address,To,CC,BCC
							$rptObj.Address = $Match.Value
							$rptObj.To = 0
							$rptObj.CC = 0
							$rptObj.BCC = 1
							$rptcntCount.Add($Match.Value,$rptObj)					
						}
					}
				}
			}
		}
		$ItemReport = "" | Select DateTimeReceived,Method,Address,Subject,Size
		$ItemReport.DateTimeReceived = $Item.DateTimeReceived
		$ItemReport.Subject = $Item.Subject
		$ItemReport.Size = $Item.Size
		$ItemReport.Method = $Method
	 	$ItemReport.Address = $Address
		$rptCollection += $ItemReport
		
    }    
    $ivItemView.Offset += $fiItems.Items.Count
	Write-Host ("Processed " + $ivItemView.Offset + " of " + $fiItems.TotalCount)
}while($fiItems.MoreAvailable -eq $true) 
$rptcntCount.Values | ft
$rptcntCount.Values | Export-Csv -NoTypeInformation -path "c:\Temp\$MailboxName-MHSummaryReport.csv"
$rptCollection | Export-Csv -NoTypeInformation -path "c:\Temp\$MailboxName-MHItemReport.csv"
