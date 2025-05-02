[CmdletBinding()]
param(
	[parameter(Mandatory=$true]
	[string]
	$Asset,
	[parameter(Mandatory=$true]
	[string]
	$Phone,
#	[parameter(Mandatory=$true]
#	[ValidateSet('MRSVR1', 'MISHERITAGE1', 'MISALMA1', 'MISBENGE1', 'MISFRISCO1', 'DC02', 'ECIRKWLDC', 'DC01')]
#	[string]
#	$Location,
	[parameter(Mandatory=$true]
	[string]
	$Subject,
	[parameter(Mandatory=$true]
	[string]
	$Description,
	[parameter(Mandatory=$true]
	[string]
	$Category,
	[parameter(Mandatory=$true]
	[string]
	$Subcategory,
	[ValidateSet('Open', 'Closed', 'On Hold']
	[string]
	$Status='Closed'
	)


$postParams = @{
				operation='AddRequest';
				requester=$Requester
				'Asset tag'=$Asset;
				'Callback phone #'=$Phone;
#				Location=$Location;
				subject=$Subject;
				description=$Description;
				Mode='Web Form'
				category=$Category;
				subcategory=$Subcategory;
				status=$Status
				technician='Wayne Reeves';
				username='wreeves';
				password=$password;
				DOMAIN_NAME='CCMHMR';
				logonDomainName='AD_AUTH'
				}
Invoke-WebRequest -Uri http://mishelp1:8080/servlets/RequestServlet -Method POST -Body $postParams
