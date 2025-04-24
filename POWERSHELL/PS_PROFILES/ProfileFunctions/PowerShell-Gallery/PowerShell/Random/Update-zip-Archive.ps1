<#
CopyLeft Svet :)
.Synopsis
   Update-Archive
.DESCRIPTION
   Long description
.SYNTAX
    Update-Archive [-nameOfFileToUpAdded] <Object> [-pathToTheCurrentArchive] <Object>  [<CommonParameters>]
.PARAMETERS
    -nameOfFileToUpAdded <Object>

        Required?                    true
        Position?                    0
        Default value                
        Accept pipeline input?       true (ByPropertyName)
        Accept wildcard characters?  

    -pathToTheCurrentArchive <Object>

        Required?                    true
        Position?                    1
        Default value                
        Accept pipeline input?       true (ByPropertyName)
.ALIASES
    update
    add
.REMARKS
    None
.EXAMPLE
    PS C:\Users\svetlozar\Desktop\0> Update-Archive
    cmdlet Update-Archive at command pipeline position 1
    Supply values for the following parameters:
    nameOfFileToUpAdded: 20-July-2022.sql
    pathToTheCurrentArchive: 20-07-2022-Wednesday.zip
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
    System.Object
.OUTPUTS
    System.Object
#>
function Update-Archive
{
    [CmdletBinding()]
    [Alias("update","add")]
   
    Param
    (
        # nameOfFileToUpAdded
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $nameOfFileToUpAdded,

        # Param2 help description
         [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $pathToTheCurrentArchive
    )

    Begin
    {
    }
    Process
    {
        Compress-Archive -Path $nameOfFileToUpAdded -Update -DestinationPath $pathToTheCurrentArchive
    }
    End
    {
    }
}

Update-Archive


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmP6oEm2//vi6IN0vA5vTIziZ
# gbigggMGMIIDAjCCAeqgAwIBAgIQEtSZ1T/kOptFn60oc6QVGTANBgkqhkiG9w0B
# AQsFADAZMRcwFQYDVQQDDA5HbyBMaXZlIFVLIEx0ZDAeFw0yMjA3MTYxMTU4MDRa
# Fw0yNzA3MTUyMTAwMDBaMBkxFzAVBgNVBAMMDkdvIExpdmUgVUsgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6sGeW+q3SYpKYZDPQu7JlI3+EWSP
# CoyXoSwzgC1SpE6vkEJYv0A4INSi4qRpiZhx7bxEREB9nDEWZTLhnXj55xe/0f6Z
# 1Xvszi9AZb0S7Hs9rIdPOYI0+FCLLauKShC7re2Gb7Yneh1GOnyLBzilj4tf2NIL
# SgaVBSS1UOH50i+MXwsBBjtGuIX8mipoGFXz2KQB7RLg0f8B4fM17X/to9/Mvh64
# LkJ641g9YbfG8hMyV49GgtOamJ8Vs0hc5M5k2T01Xen1ohDH7CT7N2MJ8DZdHWyE
# T7HCc3DFihGiELCDTL8mSeCW/9y0IZfGe/wXjswN5lumVkmgdyuAW8FEWQIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFFqWiQL2Ulc7dX4k9KFaUxF615mlMA0GCSqGSIb3DQEBCwUAA4IBAQBVBIBh
# s++k7Fe4+xqRc5N23lLDBw0+3htnzaFSxHh5geyMH9A5kDyGzcn1JpmeTyWPgdnN
# ecuL+Pu3GGPGfVP43G9Pb/E1hlh05zDbkqZK7SdxYE3hpjC+RbVoRQu0nrpDhL/x
# CNwt/WJt7eWT4Z+YRXZSPQMCa3JAMuQoZS9/q5F6npJsoErntRX//x/K6eOCff9I
# C+txUCaBrPIdjv5nuXqJb96L2zgUsUjg9XOAhPclu+Kl+NGgJF8hk+MyqB+fitkR
# pKe1GngOZqVCwTzLx56p67b7anhHn1XzNCAewcDjB3miPVR7vYN4ge1T4DBJpI5r
# bzHQtincz/LM9nLnMYIBzjCCAcoCAQEwLTAZMRcwFQYDVQQDDA5HbyBMaXZlIFVL
# IEx0ZAIQEtSZ1T/kOptFn60oc6QVGTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUQcYu8xBwgyy8
# FIUhIvXsURgMR/UwDQYJKoZIhvcNAQEBBQAEggEAD5XKf5FjuMVzlgxadRD0Z1gC
# e2sFfT0Agj0gtcgh0CMz9Mt4IsHqZ6c31CxhK060DajlYbfQ84DLHygvKR/gcei8
# C5OZbkfacyfC5jzIrM5QouaomfuxALItD+pIdbLx88BaeqvCVA/t15Ag9RjSgkdW
# fvedK/4xmbvO8i5tOAInh8W/8Yu7aZ3tEDWzbBz8E/I4QjEErkhQVh+f/eE1ipro
# mM/UFruqQprwy0rIE5Y0JlJSo/JtS2yNUxrRSZ0jVokH+Yjz2z6SiV7TLVGL1esa
# 5Kwe8i93GphQ2oqeBp7XLA4+reUKpA53ib2hkCYv3Ouz4+UnDd00BN22777fkA==
# SIG # End signature block
