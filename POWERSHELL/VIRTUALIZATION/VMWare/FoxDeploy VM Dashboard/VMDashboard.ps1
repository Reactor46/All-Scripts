Import-Module VMware.PowerCLI
Connect-VIServer -Server vvcvcsap03.vegas.com


#############
# FUNCTIONS #
#############
function Get-VMCPUAverage {
    param (
        $VM
    )
    $Start = (get-date).AddDays(-$DaysBack)
    $Finish = get-date
   

    $stats = get-vm $VM | get-stat -MaxSamples "1000" -Start $Start -Finish $Finish -Stat "cpu.usage.average" | `
    ? { ($_.TimeStamp).Hour -gt $PeakTimeStart -and ($_.TimeStamp).Hour -lt $PeakTimeEnd }
    $aggr_stats = $stats | Measure-Object -Property Value -Average
    $avg = $aggr_stats.Average
    return $avg
}


#zero out variables before beginning (helps for working on file in ISE
$main = $Null 
$i = 0 

$VMS = Get-VM | sort State 

$head = (Get-Content .\head.html) -replace '%4',(get-date).DateTime
$tail = Get-Content .\tail.html


ForEach ($VM in $VMS){
    
    #increment $i, to use the many different background image options
    $i++
    
    #set the name 
    $Name=$vm.Name
    
    #choose which style (which affect the color of the card generated)
    if ($vm.Guest.State -eq 'Off'){
        $style = 'style1'
        }
        elseif($vm.state -eq 'Running'){
        $style = "style2"
        }
        else{
        #VM is haunted or something else
        $style = "style3"
        }
    
    #if the VM is on, don't just show it's name, but it's RAM and CPU usage too
    if ($VM.Guest.State -eq 'Running'){
     
        $Name="$($VM.Name)<br>
         RAM: $($VM.MemoryMB /1mb)<br>
         CPU: $($VM.CPUUsage)"
    }

    #Set the flavor text, which appears when we hover over the card
    $description= @"
        Currently $($VM.Status.ToLower()) with a <br>
        state of $($VM.State)<br>
        It was created on $($VM.CreationTime)<br>
        Its files are found in $($VM.Path)
"@

    #make a card
    $tile = @"
                                <article class="$style">
									<span class="image">
										<img src="images/pic01.jpg" alt="" />
									</span>
									<a href="generic.html">
										<h2>$Name</h2>
										<div class="content">
											<p>$($description)</p>
										</div>
									</a>
								</article>
"@


$main += $tile
}

$html = $head + $main + $tail

$html > .\VMReport.html
