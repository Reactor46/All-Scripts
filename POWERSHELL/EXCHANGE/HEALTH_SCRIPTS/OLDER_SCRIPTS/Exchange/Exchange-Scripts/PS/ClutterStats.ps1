## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

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
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
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

function ConvertId{    
 param (
         $HexId = "$( throw 'HexId is a mandatory Parameter' )"
    )
 process{
     $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
     $aiItem.Mailbox = $MailboxName      
     $aiItem.UniqueId = $HexId   
     $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
     $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
  return $convertedId.UniqueId
 }
}

$ReportObj = "" | Select FirstMessage,LastMessage,TotalCount,TotalSize,Last7Count,Last7Size,Last7Direction,PreviousWeekCount,PreviousWeekSize,Last30Count,Last30Size,Senders,GalSenders,NumberOfAttachments,Attachments,SendersInGal,SendersNotInGal
$ReportObj.Senders = @{}
$ReportObj.GalSenders = @{}
$ReportObj.Attachments =  @{}
$ReportObj.NumberOfAttachments = 0
$ReportObj.SendersInGal = 0
$ReportObj.SendersNotInGal = 0

# Bind to the Root Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
$ClutterFolderEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([System.Guid]::Parse("{23239608-685D-4732-9C55-4C95CB4E8E33}"), "ClutterFolderEntryId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
$psPropset.Add($ClutterFolderEntryId)
$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)
$FolderIdVal = $null
if ($RootFolder.TryGetProperty($ClutterFolderEntryId,[ref]$FolderIdVal))
{
 	$Clutterfolderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((ConvertId -HexId ([System.BitConverter]::ToString($FolderIdVal).Replace("-",""))))
 	$ClutterFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Clutterfolderid);
 	#Define ItemView to retrive just 1000 Items    
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
	$fiItems = $null    
	do{ 
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
	    $fiItems = $service.FindItems($ClutterFolder.Id,$ivItemView)  
		#$fiItems = $service.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ivItemView) 
		if($fiItems.Items.Count -gt 0){
		    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
		    foreach($Item in $fiItems.Items){
				$ReportObj.TotalCount++
				if($ReportObj.TotalCount -eq 1){
					$ReportObj.LastMessage = "From : " + $Item.Sender.Name + " Subject : " + $Item.Subject
				}
				$ReportObj.FirstMessage = "From : " + $Item.Sender.Name + " Subject : " + $Item.Subject
				$ReportObj.TotalSize += $Item.Size
				##DateStuff
				if($Item.DateTimeReceived -gt (Get-Date).AddDays(-7)){
					$ReportObj.Last7Count++
					$ReportObj.Last7Size+= $Item.Size
				}
				if($Item.DateTimeReceived -gt (Get-Date).AddDays(-14) -band $Item.DateTimeReceived -le (Get-Date).AddDays(-7)){
					$ReportObj.PreviousWeekCount++
					$ReportObj.PreviousWeekSize+= $Item.Size
				}
				if($Item.DateTimeReceived -gt (Get-Date).AddMonths(-1)){
					$ReportObj.Last30Count++
					$ReportObj.Last30Size+= $Item.Size
				}
				if($Item.Attachments.Count -gt 0){
					$ReportObj.NumberOfAttachments += $Item.Attachments.Count
					foreach($Attachemnt in  $Item.Attachments){
						if(!$Attachemnt.IsInline){
							if($Attachemnt.Name -ne $null){
								if($Attachemnt.Name.Length -gt 2){
									if(!$ReportObj.Attachments.Contains($Attachemnt.Name)){
										$ReportObj.Attachments.Add($Attachemnt.Name,1)
									}
									else{
										$ReportObj.Attachments[$Attachemnt.Name]++
									}
								}
							}
						}
						
					}
				}
				if($Item.Sender -ne $null){
					if($Item.Sender.Address -ne $null){
						if(!$ReportObj.Senders.Contains($Item.Sender.Address)){
							$ReportObj.Senders.Add($Item.Sender.Address,1)
							$ncCol = $service.ResolveName($Item.Sender.Address, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $false);
							if($ncCol.Count -gt 0){
								$ReportObj.GalSenders.Add($Item.Sender.Address,1)
								$ReportObj.SendersInGal++
							}
							else{
								$ReportObj.SendersNotInGal++
							}
						}
						else{
							$ReportObj.Senders[$Item.Sender.Address] += 1
							if($ReportObj.GalSenders.Contains($Item.Sender.Address)){
								$ReportObj.GalSenders[$Item.Sender.Address] += 1
							}
						}
					}
				}
		    }
		}
	    $ivItemView.Offset += $fiItems.Items.Count    
	}while($fiItems.MoreAvailable -eq $true) 
	}
else{
 "Clutter Folder not found"
}
if($ReportObj.Last7Count -gt $ReportObj.PreviousWeekCount){
	$ReportObj.Last7Direction = "Assending"
}
else{
	$ReportObj.Last7Direction = "Descending"
}
$Reporthtml = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="content-type" content="application/xhtml+xml; charset=utf-8" />
    <meta content="width=device-width" name="viewport" />
    <title>Clutter Statistics</title>
    <meta content="width=device-width, initial-scale=1.0" name="viewport" />
    <meta content="IE=edge" http-equiv="X-UA-Compatible" />
    <meta content="telephone=no" name="format-detection" />
    <style type="text/css">  /* RESET */  #outlook a {padding:0;} body {width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0; mso-line-height-rule:exactly;}  table td { border-collapse: collapse; }  .ExternalClass {width:100%;}  .ExternalClass, .ExternalClass p, .ExternalClass span, .ExternalClass font, .ExternalClass td, .ExternalClass div {line-height: 100%;}  table td {border-collapse: collapse;}  /* IMG */  img {outline:none; text-decoration:none; -ms-interpolation-mode: bicubic;}  a img {border:none;}  /* Becoming responsive */  @media only screen and (max-device-width: 480px) {  table[id="container_div"] {max-width: 480px !important;}  table[id="container_table"], table[class="image_container"], table[class="image-group-contenitor"] {width: 100% !important; min-width: 320px !important;}  table[class="image-group-contenitor"] td, table[class="mixed"] td, td[class="mix_image"], td[class="mix_text"], td[class="table-separator"], td[class="section_block"] {display: block !important;width:100% !important;}  table[class="image_container"] img, td[class="mix_image"] img, table[class="image-group-contenitor"] img {width: 100% !important;}  table[class="image_container"] img[class="natural-width"], td[class="mix_image"] img[class="natural-width"], table[class="image-group-contenitor"] img[class="natural-width"] {width: auto !important;}  a[class="button-link justify"] {display: block !important;width:auto !important;}  td[class="table-separator"] br {display: none;}  td[class="cloned_td"]  table[class="image_container"] {width: 100% !important; min-width: 0 !important;} } table[class="social_wrapp"] {width: auto;} </style>
  </head>
  <body style="    background-color: #d5e4ed;">
    <table width="100%" cellspacing="0" cellpadding="0" border="0" bgcolor="#d5e4ed"

      align="center" style="text-align:center; background-color:#d5e4ed; border-collapse: collapse"

      id="container_div">
      <tbody>
        <tr>
          <td align="center"><br />
            <table cellspacing="0" cellpadding="0" border="0" id="container_wrapper">
              <tbody>
                <tr>
                  <td>
                    <table width="600" cellspacing="0" cellpadding="0" border="0"

                      bgcolor="#ffffff" style="border-collapse: collapse; min-width: 600px;"

                      id="container_table">
                      <tbody>
                        <tr>
                          <td valign="top" bgcolor="#ffffff">
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#007fff" style="border-collapse: collapse; background-color: rgb(0, 127, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(255, 255, 255);">
                                    <div style="text-align: center; color: rgb(255, 255, 255);"><span

                                        style="color: rgb(255, 255, 255); line-height: 130%;"><span

                                          style="font-size: 28px; line-height: 130%;"><span

                                            style="font-family: tahoma, geneva, sans-serif; line-height: 130%;"><strong

                                              style="color: rgb(255, 255, 255);">Clutter
                                              Statistics</strong></span></span></span></div>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                        <tr>
                          <td valign="top" bgcolor="#ffffff">
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#00a8cc" style="border-collapse: collapse; background-color: rgb(0, 168, 204);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);">
                                    <table cellspacing="0" cellpadding="5" border="0"

                                      align="center" class="cke_show_border">
                                      <tbody>
                                        <tr>
                                          <td colspan="2" rowspan="1" style="width: 600px; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><u><strong><span

                                                  style="font-size: 26px; line-height: 130%;">Folder
                                                  Totals</span></strong></u></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#1#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);">
                                            <div><span style="font-size: 18px; line-height: 130%;"><strong>Total
                                                  Item Count</strong></span></div>
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#2#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Item Size (MB)</strong></span></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                        <tr>
                          <td valign="top" bgcolor="#ffffff">
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#00a8cc" style="border-collapse: collapse; background-color: rgb(0, 168, 204);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: rgb(0, 0, 0); margin: 0px auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#6fcee2" style="border-collapse: collapse; background-color: rgb(111, 206, 226);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);">
                                    <table cellspacing="0" cellpadding="5" border="0"

                                      align="center" class="cke_show_border">
                                      <tbody>
                                        <tr>
                                          <td colspan="2" rowspan="1" style="width: 600px; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><strong><u><span

                                                  style="font-size: 26px; line-height: 130%;">Last
                                                  7 Days</span></u></strong></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#3#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Item Count</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#4#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Tota
                                                Item Size (MB)</strong></span></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#00d2ff" style="border-collapse: collapse; background-color: rgb(0, 210, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: rgb(0, 0, 0); margin: 0px auto; color: rgb(0, 0, 0);"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#00d2ff" style="border-collapse: collapse; background-color: rgb(0, 210, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);"><span

                                      style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: rgb(102, 102, 102); line-height: 130%;"></span>
                                    <table cellspacing="0" cellpadding="5" border="0"

                                      align="center" class="cke_show_border">
                                      <tbody>
                                        <tr>
                                          <td colspan="2" rowspan="1" style="width: 600px; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><u><strong><span

                                                  style="font-size: 26px; line-height: 130%;">7-14
                                                  Days</span></strong></u></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#5#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Item Count</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#6#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Tota
                                                Item Size (MB)</strong></span></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                    <span style="font-size: 14px; font-family: Arial, Helvetica, sans-serif; color: rgb(102, 102, 102); line-height: 130%;">
                                    </span></td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#00d2ff" style="border-collapse: collapse; background-color: rgb(0, 210, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: #333; margin: 0 auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#66e4ff" style="border-collapse: collapse; background-color: rgb(102, 228, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);">
                                    <table cellspacing="0" cellpadding="5" border="0"

                                      align="center" class="cke_show_border">
                                      <tbody>
                                        <tr>
                                          <td colspan="2" rowspan="1" style="width: 600px; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><u><strong><span

                                                  style="font-size: 26px; line-height: 130%;">Last
                                                  30 Days</span></strong></u></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#7#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Item Count</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#8#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Tota
                                                Item Size (MB)</strong></span></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#66e4ff" style="border-collapse: collapse; background-color: rgb(102, 228, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: #333; margin: 0 auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#93ecff" style="border-collapse: collapse; background-color: rgb(147, 236, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);">
                                    <table cellspacing="0" cellpadding="5" border="0"

                                      align="center" class="cke_show_border">
                                      <tbody>
                                        <tr>
                                          <td colspan="2" rowspan="1" style="width: 600px; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><u><strong><span

                                                  style="font-size: 26px; line-height: 130%;">Folder
                                                  Statistics</span></strong></u></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><br />
                                          </td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#9#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Number of Attachments</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#10#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Total
                                                Number of Senders</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="width: 20%; text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#11#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Senders
                                                in Gal</strong></span></td>
                                        </tr>
                                        <tr>
                                          <td style="text-align: right; text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 20px; line-height: 130%;">#12#</span></td>
                                          <td style="text-decoration: none; padding: 5px; line-height: 130%; color: rgb(0, 0, 0);"><span

                                              style="font-size: 18px; line-height: 130%;"><strong>Senders
                                                not in Gal</strong></span></td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#93ecff" style="border-collapse: collapse; background-color: rgb(147, 236, 255);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: #333; margin: 0 auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#b2e5e5" style="border-collapse: collapse; background-color: rgb(178, 229, 229);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);"><span

                                      style="font-size: 22px; line-height: 130%;"><b><u><span

                                            style="font-size: 26px; line-height: 130%;">Top
                                            Senders</span></u><br />
                                        <br />
                                      </b><span style="font-size: 18px; line-height: 130%;">#13#</span></span></td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#b2e5e5" style="border-collapse: collapse; background-color: rgb(178, 229, 229);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: #333; margin: 0 auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#a0cece" style="border-collapse: collapse; background-color: rgb(160, 206, 206);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);"><strong><u><span

                                          style="font-size: 26px; line-height: 130%;">Top
                                          Gal Senders</span></u></strong><br />
                                    <br />
                                    <span style="font-size: 18px; line-height: 130%;">#14#</span></td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#a0cece" style="border-collapse: collapse; background-color: rgb(160, 206, 206);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" height="2" align="left">
                                    <table width="100%" cellspacing="0" cellpadding="0"

                                      border="0" align="center" style="border-top-width: 1px; border-top-style: solid; border-top-color: #333; margin: 0 auto;"

                                      class="divider">
                                      <tbody>
                                        <tr>
                                          <td style="line-height: 0; text-decoration: none; padding: 0px;">
                                              </td>
                                        </tr>
                                      </tbody>
                                    </table>
                                  </td>
                                </tr>
                              </tbody>
                            </table>
                            <table width="100%" cellspacing="0" cellpadding="10"

                              border="0" bgcolor="#86becb" style="border-collapse: collapse; background-color: rgb(134, 190, 203);">
                              <tbody>
                                <tr valign="top">
                                  <td valign="top" style="line-height: 130%; color: rgb(0, 0, 0);"><span

                                      style="font-size: 26px; line-height: 130%;"><strong></strong><u><strong>Top
                                          Attachment Names​</strong><br />
                                        </u><span style="font-size: 18px; line-height: 130%;">#15#</span></span></td>
                                </tr>
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table>
  </body>
</html>
		
"@
$ReportObj
$Reporthtml = $Reporthtml.replace("#1#",$ReportObj.TotalCount)
$Reporthtml = $Reporthtml.replace("#2#",[Math]::Round(($ReportObj.TotalSize/1024/1024),2))
$Reporthtml = $Reporthtml.replace("#3#",$ReportObj.Last7Count)
$Reporthtml = $Reporthtml.replace("#4#",[Math]::Round(($ReportObj.Last7Size/1024/1024),2))
$Reporthtml = $Reporthtml.replace("#5#",$ReportObj.PreviousWeekCount)
$Reporthtml = $Reporthtml.replace("#6#",[Math]::Round(($ReportObj.PreviousWeekSize/1024/1024),2))
$Reporthtml = $Reporthtml.replace("#7#",$ReportObj.Last30Count)
$Reporthtml = $Reporthtml.replace("#8#",[Math]::Round(($ReportObj.Last30Size/1024/1024),2))
$Reporthtml = $Reporthtml.replace("#9#",$ReportObj.NumberOfAttachments)
$Reporthtml = $Reporthtml.replace("#10#",$ReportObj.Senders.Count)
$Reporthtml = $Reporthtml.replace("#11#",$ReportObj.SendersInGal)
$Reporthtml = $Reporthtml.replace("#12#",$ReportObj.SendersNotInGal)
$tc = 0
$ReportObj.Senders.GetEnumerator() | Sort-Object Value -descending | ForEach-Object {
	$tc++
	if($tc -le 10){
		$SendersList = $SendersList + $_.Name + " : " + $_.Value + "<br>"
	}
}
$tc = 0
$ReportObj.GalSenders.GetEnumerator() | Sort-Object Value -descending | ForEach-Object {
	$tc++
	if($tc -le 10){
		$GalSendersList = $GalSendersList + $_.Name + " : " + $_.Value + "<br>"
	}
}
$tc = 0
$ReportObj.Attachments.GetEnumerator() | Sort-Object Value -descending | ForEach-Object {
	$tc++
	if($tc -le 10){
		$AttachmentsList = $AttachmentsList + $_.Name + " : " + $_.Value + "<br>"
	}
}
$Reporthtml = $Reporthtml.replace("#13#",$SendersList)
$Reporthtml = $Reporthtml.replace("#14#",$GalSendersList)
$Reporthtml = $Reporthtml.replace("#15#",$AttachmentsList)

$toAddress = $MailboxName
$EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
$EmailMessage.Subject = "Clutter Statistics"  
#Add Recipients    
$EmailMessage.ToRecipients.Add($toAddress)  
$EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody  
$EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML  
$EmailMessage.Body.Text = $Reporthtml
$EmailMessage.Send()  