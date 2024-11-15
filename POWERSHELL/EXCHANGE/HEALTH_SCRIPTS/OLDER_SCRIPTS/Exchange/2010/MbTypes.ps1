####################### 
<# 
.SYNOPSIS 
  Performs an eDiscovery Search using Powershell and the Exchange Web Services API in a Mailbox and produces a CSV report of ItemType in that Mailbox
 
.DESCRIPTION 
  Performs an eDiscovery Search using Powershell and the Exchange Web Services API in a Mailbox and produces a CSV report of ItemType in that Mailbox
 
 Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951

.EXAMPLE
 PS C:\>Get-MailboxItemTypeStats -MailboxName user.name@domain.com -OutputFileName Report.csv

 This Example Performs an eDiscovery Search using Powershell and the Exchange Web Services API in a Mailbox and produces a CSV report of ItemType in that Mailbox
#> 
function Get-MailboxItemTypeStats 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
		[Parameter(Position=2, Mandatory=$true)] [string]$OutputFileName
    )  
 	Begin
		 {
		$KQL = "kind:email OR kind:meetings OR kind:contacts OR kind:tasks OR kind:notes OR kind:IM OR kind:rssfeeds OR kind:voicemail";			 
		$SearchableMailboxString = $MailboxName;

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
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1  
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
		$service.Credentials = $creds      
		#$service.TraceEnabled = $true
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

		$gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
		$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
		$mbCount = 0;
		foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
		{
		    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
		    $mbCount++;
		}
		$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
		$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
		$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
		$mbqa[0] = $mbq
		$smSearchMailbox.SearchQueries = $mbqa;
		$smSearchMailbox.PageSize = 100;
		$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
		$smSearchMailbox.PerformDeduplication = $false;           
		$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::StatisticsOnly;
		$srCol = $service.SearchMailboxes($smSearchMailbox);
		$rptCollection = @()

		if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
		{
			foreach($KeyWorkdStat in $srCol[0].SearchResult.KeywordStats){
				if($KeyWorkdStat.Keyword.Contains(" OR ") -eq $false){
					$rptObj = "" | Select ItemType,ItemHits,Size
					$rptObj.ItemType = $KeyWorkdStat.Keyword.Replace("Kind:","")
					$rptObj.ItemHits = $KeyWorkdStat.ItemHits
					$rptObj.Size = [System.Math]::Round($KeyWorkdStat.Size /1024/1024,2)
					$rptCollection += $rptObj
				}
			}   
		}
		Write-Output $rptCollection
		$rptCollection | Export-Csv -NoTypeInformation -Path $OutputFileName

		}
}