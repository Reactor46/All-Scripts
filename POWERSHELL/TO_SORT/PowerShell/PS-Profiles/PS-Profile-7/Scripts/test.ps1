        [CmdletBinding()]
        Param(
            # Any other parameters can go here
        )
    
        DynamicParam 
                {
                # Set the dynamic parameters' name. You probably want to change this.
                $ParameterName = 'Server'
    
                # Create the dictionary 
                $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
    
                # Create the collection of attributes
                $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
    
                # Create and set the parameters' attributes. You may also want to change these.
                $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
                #$ParameterAttribute.Mandatory = $true
                #$ParameterAttribute.Position = 1
    
                # Add the attributes to the attributes collection
                $AttributeCollection.Add($ParameterAttribute)
    
                # Generate and set the ValidateSet. You definitely want to change this. This part populates your set. 
                Get-ADDomainController -Filter { OperatingSystem -notlike "Windows Server 2003" -and OperatingSystem -notlike "Windows Server® 2008 Standard" }  | ForEach-Object { [array]$arrset += $_.name}
                $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)
    
                # Add the ValidateSet to the attributes collection
                $AttributeCollection.Add($ValidateSetAttribute)
    
                # Create and return the dynamic parameter
                $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
                $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
                return $RuntimeParameterDictionary
                }       
    
        begin 
                {
                # Bind the parameter to a friendly variable
                $Server = $PsBoundParameters[$ParameterName]
                }       
    
        process 
                {
                # Your code goes here
                $dcs = $RuntimeParameter.Attributes.ValidValues
                $dcs
                }       
