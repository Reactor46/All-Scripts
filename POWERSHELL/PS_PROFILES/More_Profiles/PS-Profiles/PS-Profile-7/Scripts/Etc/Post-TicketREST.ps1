[CmdletBinding()]
param(
    $requester,
    $tech
    )
if ( !$tech ) 
    {
    Write-Host `
    "
    1. Connie
    2. Chad
    3. Melissa
    4. Wayne
    " -ForegroundColor Cyan
    $tech = Read-Host "Technician"
    switch ( $tech )
        {
        1 { $tech = "Connie" }
        2 { $tech = "Chad" }
        3 { $tech = "Melissa" }
        4 { $tech = "Wayne" }
        }
    }
switch ($tech) 
    { 
    "Chad" 
        {  
        $techname = "Chad Hickerson"
        $techkey = '3F40C808-FD88-49CC-A1CC-D027390976A6'
        }
    "Connie" 
        {
        $techname = "Connie Wilson"
        $techkey = '380D59DA-17EF-4341-9397-37319237D7AA'
        }
    "Melissa"
        {
        $techname = "Melissa Bolton"
        $techkey = '3079B86E-D5FC-40D2-AFA8-850A9F70A959'
        }
    "Wayne" 
        {
        $techname = "Wayne Reeves"
        $techkey = "92741403-5BF7-413B-998D-6F354C1DD900"
        }
    }
Write-Host $techname
$inputData = @"
<Operation>
<Details>
<requester>$requester</requester>
<requesttemplate>Account Lockout - Network</requesttemplate>
<priority>Normal</priority>
<location>BH-McKinney-Crisis Center-1416 N Church</location>
<On-call>Y</On-call>
<technician>$techname</technician>
<status>Closed</status>
</Details>
</Operation>
"@

$postParams = @{OPERATION_NAME='ADD_REQUEST';TECHNICIAN_KEY=$techkey;INPUT_DATA=$inputData}

#$results = Invoke-WebRequest -Uri http://mishelp1:8080/sdpapi/request -Method POST -Body $postParams
#$results.content