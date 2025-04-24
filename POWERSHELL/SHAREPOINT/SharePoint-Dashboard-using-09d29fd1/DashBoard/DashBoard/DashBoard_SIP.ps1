###############################################################################
#             Authors: Vikas Sukhija
#             Date: 08/12/2013
#             Description: Script to update sharepoint lists on SIP site
#                          This will enable updating of DashBoard 
###############################################################################

If ((Get-PSSnapin | where {$_.Name -match "SharePoint.Powershell"}) -eq $null)
{
	Add-PSSnapin Microsoft.SharePoint.Powershell
}


$date = get-date -format d
$date = $date.ToString().Replace(“/”, “-”)

$siteurl = "http://sharepointlab/sites/SIP"
$ListName = "SIPs" 

##################################Get List##########################

$spWeb = Get-SPWeb $siteurl                                                  
$list = $spWeb.Lists[$ListName]
$items = $list.getitems()


####################################Create Count variables############
$Collection=@()


#####################Total SIPs######################################
function TotalSips ($IsTower) {


foreach($item in $items)

{

if($item["IS Tower"] -eq $IsTower)
{
$Collection += $item.Name

}

}
Write-host $Collection.count
$output = ".\Dashboard\" + $IsTower + "-" + "Totsip" + "_.txt"
$Collection.count | out-file $output

}

###########################Sips Closed by Tower##########################

function sipsclosedbytower ($IsTower) {


foreach($item in $items)

{

if(($item["IS Tower"] -eq $IsTower) -and ($item["SIP Status"] -eq "Closed"))
{
$Collection += $item.Name

}

}
Write-host $Collection.count
$output = ".\Dashboard\" + $IsTower + "-" + "sipclbt" + "_.txt"
$Collection.count | out-file $output

}

###########################Sips in-progress by Tower##########################

function sipsinprogressbytower ($IsTower) {


foreach($item in $items)

{

if(($item["IS Tower"] -eq $IsTower) -and ($item["SIP Status"] -eq "In Progress"))
{
$Collection += $item.Name

}

}
Write-host $Collection.count
$output = ".\Dashboard\" + $IsTower + "-" + "sipinprgt" + "_.txt"
$Collection.count | out-file $output
}

################################Total Active,Closed,progress,Proposed,Rejected######

function sipstatus ($SIPSTATUS) {


foreach($item in $items)

{

if($item["SIP Status"] -eq $SIPSTATUS)
{
$Collection += $item.Name

}

}
Write-host $Collection.count
$output = ".\Dashboard\" + $SIPSTATUS + "-" + "sipst" + "_.txt"
$Collection.count | out-file $output

}

####################################################################################

TotalSips Messaging
TotalSips Windows
TotalSips Unix
TotalSips DBA
TotalSips DCO
TotalSips Storage
TotalSips Montioring
TotalSips Tools
TotalSips Voice
TotalSips Network
TotalSips Security
TotalSips Helpdesk
TotalSips Remedy
TotalSips Process
TotalSips WSM

sipsclosedbytower Messaging
sipsclosedbytower Windows
sipsclosedbytower Unix
sipsclosedbytower DBA
sipsclosedbytower DCO
sipsclosedbytower Storage
sipsclosedbytower Montioring
sipsclosedbytower Tools
sipsclosedbytower Voice
sipsclosedbytower Network
sipsclosedbytower Security
sipsclosedbytower Helpdesk
sipsclosedbytower Remedy
sipsclosedbytower Process
sipsclosedbytower WSM

sipsinprogressbytower Messaging
sipsinprogressbytower Windows
sipsinprogressbytower Unix
sipsinprogressbytower DBA
sipsinprogressbytower DCO
sipsinprogressbytower Storage
sipsinprogressbytower Montioring
sipsinprogressbytower Tools
sipsinprogressbytower Voice
sipsinprogressbytower Network
sipsinprogressbytower Security
sipsinprogressbytower Helpdesk
sipsinprogressbytower Remedy
sipsinprogressbytower Process
sipsinprogressbytower WSM

sipstatus Active
sipstatus Closed
sipstatus "In Progress"
sipstatus Proposed
sipstatus Rejected
#####################################Update Towers List#################################

Function dashupdatelist ($ISTower)
{

$ListName = "Towers"
$list = $spWeb.Lists[$ListName]
$items = $list.getitems()

$input1 = ".\Dashboard\" + "$ISTower" + "-" + "Totsip" + "_.txt"
$value = get-content $input1
write-host $value

foreach($item in $items)
{
if($item["IS Tower"] -eq $ISTower)

{
if($item["TotalSips"] -ne $value)
{

$item["TotalSips"] = $value;
$item.update();

}

}
}
}

#####################################Update sip closed by towers List##########################

Function dashupdatesipcltlist ($ISTower)
{

$ListName = "SIPClosedTower"
$list = $spWeb.Lists[$ListName]
$items = $list.getitems()

$input1 =  ".\Dashboard\" + $ISTower + "-" + "sipclbt" + "_.txt"
$value = get-content $input1
write-host $value

foreach($item in $items)
{
if($item["IS Tower"] -eq $ISTower)

{
if($item["Count"] -ne $value)
{

$item["Count"] = $value;
$item.update();

}

}
}
}

#####################################Update sip inprogress by towers List##########################

Function dashupdatesipinpltlist ($ISTower)
{

$ListName = "SipsInProgressbyTowers"
$list = $spWeb.Lists[$ListName]
$items = $list.getitems()

$input1 =  ".\Dashboard\" + $ISTower + "-" + "sipinprgt" + "_.txt"
$value = get-content $input1
write-host $value

foreach($item in $items)
{
if($item["IS Tower"] -eq $ISTower)

{
if($item["Count"] -ne $value)
{

$item["Count"] = $value;
$item.update();

}

}
}
}

#####################################Update sip status List##########################

Function dashupdatesipstsltlist ($SIPSTATUS)
{

$ListName = "InProgress"
$list = $spWeb.Lists[$ListName]
$items = $list.getitems()

$input1 = ".\Dashboard\" + $SIPSTATUS + "-" + "sipst" + "_.txt"
$value = get-content $input1
write-host $value

foreach($item in $items)
{
if($item["Sip Status"] -eq $SIPSTATUS)

{
if($item["Count"] -ne $value)
{

$item["Count"] = $value;
$item.update();

}

}
}
}

#########################################Call Function##########################################
dashupdatelist Messaging
dashupdatelist Windows
dashupdatelist Unix
dashupdatelist DBA
dashupdatelist DCO
dashupdatelist Storage
dashupdatelist Montioring
dashupdatelist Tools
dashupdatelist Voice
dashupdatelist Network
dashupdatelist Security
dashupdatelist Helpdesk
dashupdatelist Remedy
dashupdatelist Process
dashupdatelist WSM

dashupdatesipcltlist Messaging
dashupdatesipcltlist Windows
dashupdatesipcltlist Unix
dashupdatesipcltlist DBA
dashupdatesipcltlist DCO
dashupdatesipcltlist Storage
dashupdatesipcltlist Montioring
dashupdatesipcltlist Tools
dashupdatesipcltlist Voice
dashupdatesipcltlist Network
dashupdatesipcltlist Security
dashupdatesipcltlist Helpdesk
dashupdatesipcltlist Remedy
dashupdatesipcltlist Process
dashupdatesipcltlist WSM

dashupdatesipinpltlist Messaging
dashupdatesipinpltlist Windows
dashupdatesipinpltlist Unix
dashupdatesipinpltlist DBA
dashupdatesipinpltlist DCO
dashupdatesipinpltlist Storage
dashupdatesipinpltlist Montioring
dashupdatesipinpltlist Tools
dashupdatesipinpltlist Voice
dashupdatesipinpltlist Network
dashupdatesipinpltlist Security
dashupdatesipinpltlist Helpdesk
dashupdatesipinpltlist Remedy
dashupdatesipinpltlist Process
dashupdatesipinpltlist WSM

dashupdatesipstsltlist Active
dashupdatesipstsltlist Closed
dashupdatesipstsltlist "In Progress"
dashupdatesipstsltlist Proposed
dashupdatesipstsltlist Rejected

###########################################################################################################