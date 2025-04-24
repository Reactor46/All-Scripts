filter Set-Type {
<#
.SYNOPSIS
    Sets the data type of a property given the property name and the data type.
.DESCRIPTION
    Sets the data type of a property given the property name and the data type.
    This is needed as cmdlets such as Import-CSV pulls everything in as a string
    datatype so you can't sort numerically or date wise.
.PARAMETER TypeHash
    A hashtable of property names and their associated datatype
.NOTES
    # inspired by https://mjolinor.wordpress.com/2011/05/01/typecasting-imported-csv-data/

    Changes
    * reworked with begin, process, end blocks
    * reworked logic to work properly with pwsh and powershell
.EXAMPLE
    $csv = Import-CSV -Path .\test.csv | Set-Type -TypeHash @{ 'LastWriteTime' = 'DateTime'}
.LINK
    about_Properties
#>

    [cmdletbinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    param(
        [parameter(ValueFromPipeLine)]
        [psobject[]] $InputObject,

        [hashtable] $TypeHash
    )

    begin {
        Write-Verbose -Message "Starting [$($MyInvocation.Mycommand)]"
    }

    process {
        foreach ($curObject in $InputObject) {
            foreach ($key in $($TypeHash.keys)) {
                $curObject.$key = $($curObject.$key -as $($TypeHash[$key]))
            }
            $curObject
        }
    }

    end {
        Write-Verbose -Message "Ending [$($MyInvocation.Mycommand)]"
    }
}
