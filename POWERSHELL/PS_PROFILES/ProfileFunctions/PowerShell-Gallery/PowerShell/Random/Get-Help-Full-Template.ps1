
NAME
    Get-Help
    
SYNTAX
    Get-Help [[-Name] <string>] [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Full] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    Get-Help [[-Name] <string>] -Detailed [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    Get-Help [[-Name] <string>] -Examples [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    Get-Help [[-Name] <string>] -Parameter <string[]> [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    Get-Help [[-Name] <string>] -Online [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    Get-Help [[-Name] <string>] -ShowWindow [-Path <string>] [-Category {Alias | Cmdlet | Provider | General | FAQ | Glossary | HelpFile | ScriptCommand | Function | Filter | ExternalScript | All | DefaultHelp | DscResource | Class | Configuration}] [-Component <string[]>] [-Functionality <string[]>] [-Role <string[]>] [<CommonParameters>]
    
    
PARAMETERS
    -Category <string[]>
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Component <string[]>
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Detailed
        
        Required?                    true
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           DetailedView
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Examples
        
        Required?                    true
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           Examples
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Full
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           AllUsersView
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Functionality <string[]>
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Name <string>
        
        Required?                    false
        Position?                    0
        Accept pipeline input?       true (ByPropertyName)
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Online
        
        Required?                    true
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           Online
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Parameter <string[]>
        
        Required?                    true
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           Parameters
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Path <string>
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -Role <string[]>
        
        Required?                    false
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           (All)
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    -ShowWindow
        
        Required?                    true
        Position?                    Named
        Accept pipeline input?       false
        Parameter set name           ShowWindow
        Aliases                      None
        Dynamic?                     false
        Accept wildcard characters?  false
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https://go.microsoft.com/fwlink/?LinkID=113216). 
    
    
INPUTS
    System.String
    
    
OUTPUTS
    System.Object
    
ALIASES
    None
    

REMARKS
    Get-Help cannot find the Help files for this cmdlet on this computer. It is displaying only partial help.
        -- To download and install Help files for the module that includes this cmdlet, use Update-Help.
        -- To view the Help topic for this cmdlet online, type: "Get-Help Get-Help -Online" or
           go to https://go.microsoft.com/fwlink/?LinkID=2096483.



# SIG # Begin signature block
# MIIFagYJKoZIhvcNAQcCoIIFWzCCBVcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUiAGM9AjBI/qWF+Lmuy3sL/lg
# 2UegggMGMIIDAjCCAeqgAwIBAgIQJLzUFQ0UMqZIQD0LynYJzzANBgkqhkiG9w0B
# AQsFADAZMRcwFQYDVQQDDA5HbyBMaXZlIFVLIEx0ZDAeFw0yMjA3MTYxMTU3MzZa
# Fw0yNzA3MTUyMTAwMDBaMBkxFzAVBgNVBAMMDkdvIExpdmUgVUsgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jBbQhei1OZxT2X07woTaSXAAIhc
# nY1a8OwQdzKLpWKCskGbDmGBH0bdqLq1m+y8jWKnf1tEnzaFVxUKrcwq57hy8og2
# xK9Ciq2+X43ZpA4acLUnvK0auUNig2vwWXlnPr4vh4g6+E6nMvvyhPA5ND+Uvhm2
# wSi7Z+qlmYQqm7wI7IEoaGy/8zyPB1vnWOL2UUYL8oTSklGdA6/wPGcV0S41vfGR
# 1YClE/AaAMB2+GZgkGgw9gQvKwsNRFIJSKu3OHICfrPRJgObdMnKr6F2jFl2LP+E
# Rf/Qy7BkTMwLIuxm2Tik+yPh3ZLGWGr2TtQoipNcdIg/7jmGuepTssD5/QIDAQAB
# o0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFDBp69/PtsKtexPQHwwNhwtX73JqMA0GCSqGSIb3DQEBCwUAA4IBAQBQ4+Qm
# FQd/u421ZKpspyhsYw/Hjnn27uBwJ1420lpQSApuI6BnsmlZlpNVAl5ETMtbzezo
# x4xwkQ2H4/+FWg4/Mic0adMwAgqGEeOpWYZxpUQFnzCIlyH6QmXvDLuR5FwVUFKt
# fdbUj5waOOY6UJ3MSGWEykAX3R3jF/g9AgZmfvTeK9ko72QU595uIHRRXubq3akk
# 8WGiZcvgZ6WwzbHDw3iM3Q16OKOGKC/UQgJ1eiawavJUF7DQawB1VonfnbT9caZB
# 8H88Qf8u3g2FQlStoZ1lYfIUj37PZfbaBhVVZyMS7FxmSmfLXvQLtV0pT7nSRK63
# Ki30MG53JtJQdqYmMYIBzjCCAcoCAQEwLTAZMRcwFQYDVQQDDA5HbyBMaXZlIFVL
# IEx0ZAIQJLzUFQ0UMqZIQD0LynYJzzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUvHonm7T7aFZ8
# z7QopjWFQZMs0CAwDQYJKoZIhvcNAQEBBQAEggEAQqUyD0Zh19iw2gXHGIXr0Onr
# 5UxJBtgih780USsROSCJaffz9oj053zk9d77/MBjd8fzv9Xoh8i/ylwImN9VA7nI
# 3ZJyvbFDUSIgjJKx5WS/dlY1F5T6kUMHNu2RmLOh4ccJFjLc9szGSUvrECEAe5RM
# ETAGz7FhkER4EUnk/AE3gguqTCPVIupfUrwDIspmj9XtziNX+Wb24us3cdMmpE5R
# oM3Hgj4YMZfcYTEXsY4EoxczYNi+mQFOzmU8P4+bhXxemN6N5GJbMIh4jF9ygPf3
# rdvXasryefvL8ZP1VVw6Yrof6BXFp+m8CUkuw48vrgyB2pb5QeG8Q7z80KNinw==
# SIG # End signature block
