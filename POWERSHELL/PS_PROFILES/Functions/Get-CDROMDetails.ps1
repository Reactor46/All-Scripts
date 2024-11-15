function Get-CDROMDetails {                        
[cmdletbinding()]                        
param(                        
    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]                        
    [string[]]$ComputerName = $env:COMPUTERNAME                        
)                        
            
begin {}                        
process {                        
    foreach($Computer in $COmputerName) {                        
    $object = New-Object –TypeName PSObject –Prop(@{                        
                'ComputerName'=$Computer.ToUpper();                        
                'CDROMDrive'= $null;                        
                'Manufacturer'=$null                        
               })                        
    if(!(Test-Connection -ComputerName $Computer -Count 1 -ErrorAction 0)) {                        
        Write-Verbose "$Computer is OFFLINE"                        
    }                        
    try {                        
        $cd = Get-WMIObject -Class Win32_CDROMDrive -ComputerName $Computer -ErrorAction Stop                        
    } catch {                        
        Write-Verbose "Failed to Query WMI Class"                        
        Continue;                        
    }                        
            
    $Object.CDROMDrive = $cd.Drive                        
    $Object.Manufacturer = $cd.caption                        
    $Object                           
            
    }                        
}                        
            
end {}               
}