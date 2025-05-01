<#
.Synopsis
    Create-BasicTask cmdlet registers a scheduled task definition on a local computer
.DESCRIPTION
   Creating basic scheduled task in Windows - this will not configure it for Windows 10
.EXAMPLE
    TaskPath TaskName              State     
    -------- --------              -----     
    \        Win32_LogicalDisk     Ready
.SYNTAX
   Create-BasicTask [[-execute] <string>] [[-argument] <string>] [[-MultipleInstances] <string>] [[-description] <string>] [[-taskname] <string>] [<CommonParameters>]]
.PARAMETERS
    -LogonType <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -MultipleInstances <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -RunLevel <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -argument <string>
        
        Required?                    false
        Position?                    1
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -description <string>
        
        Required?                    false
        Position?                    3
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -execute <string>
        
        Required?                    false
        Position?                    0
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -taskname <string>
        
        Required?                    false
        Position?                    4
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -time <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        
    -userID <string>
        
        Required?                    false
        Position?                    2
        Accept pipeline input?       true (ByValue, ByPropertyName, FromRemainingArguments)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false  

        
        <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see 
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).
         
.INPUTS
    System.String 
.OUTPUTS
    System.Object
.NOTES
   
.COMPONENT
   
.ROLE
   
.FUNCTIONALITY
   
.Aliases
    cbs
    Basic-Task
    New-BasicTask
.REMARKS
    None
.LINKS
    None
#>
function Create-BasicTask{


    [CmdletBinding(
        ConfirmImpact='Medium')]

    [Alias("cbs","Basic-Task", "New-BasicTask")]

    Param
    (
       [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$execute="Powershell.exe",

       [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=1)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]                  
       [string]$argument="C:\scripts\Win32_LogicalDisk.ps1",

       [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=2)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$time="9PM" ,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=3)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$userID="NT AUTHORITY\SYSTEM" ,

       [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=4)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$LogonType="ServiceAccount" ,

       [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=2)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$RunLevel="Highest" ,

         [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=5)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$MultipleInstances="Parallel" ,

       [Parameter(Mandatory=$false,
                    ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=6)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$description="TESTWin32_LogicalDisk",
       
      [Parameter(Mandatory=$false,
                    ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$true,
                   Position=7)]
       [ValidateNotNull()]
       [ValidateNotNullOrEmpty()]             
       [string]$taskname=$description          
   )      
    

    Begin
    {
        $action = New-ScheduledTaskAction -Execute $execute -Argument $argument
        $trigger = New-ScheduledTaskTrigger -Daily -At $time
        $principal = New-ScheduledTaskPrincipal -UserID $userID -LogonType $LogonType -RunLevel $RunLevel
        $settings = New-ScheduledTaskSettingsSet -MultipleInstances $MultipleInstances
        Register-ScheduledTask -Action $action -Description $description -Trigger $trigger -TaskName $taskname -Settings $settings -Principal $principal
    }
    Process
    {
       
    }
    End
    {
    }   
   
}


Create-BasicTask


# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYwDy98bdery1Z9Rw3IcE842x
# UwSgggMGMIIDAjCCAeqgAwIBAgIQEtSZ1T/kOptFn60oc6QVGTANBgkqhkiG9w0B
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUHpJV5YWF4EYW
# PW1e52mhxfOGsHAwDQYJKoZIhvcNAQEBBQAEggEAQ/AzdvJF9CXSMbcnjZOvNhSC
# eNnvIj5fJFp0s/6bqVrxRtwHMh4pRRB19sytK4i5YsjQe+7XKLDPcSZ3hxFIuE6r
# ijm2DOaUzcRUFa6EMmf8rhntLyy+ONhRY4aE9RKGNOlyhNrrhx/GjIJvkkKWe/Xa
# XlTiH2rRnftW8hEwD4sIIhby3GOjZdm8w+C1DeQ3qpRA7O41rmQhjaDqCmfGRPuI
# L4tBSygzEysDmwTDkN5Tk3lrF8jjQzgMaO7nw+KIsU4ONYRwJbC+CaGr1yCUg49R
# ll8OrtuX2RyFLQ31KwW4nRD8ggBtTf0WEyL92ODvb6umYsaRxpSZwGXNCWQQzQ==
# SIG # End signature block
