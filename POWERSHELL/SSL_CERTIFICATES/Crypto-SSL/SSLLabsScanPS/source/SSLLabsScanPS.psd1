@{
    # Script module or binary module file associated with this manifest.
    RootModule        = 'SSLLabsScanPS.psm1'

    # Version number of this module.
    ModuleVersion     = '0.0.1'

    # ID used to uniquely identify this module
    GUID              = 'a0b15365-92e6-48f9-8345-2346931ed2d0'

    # Author of this module
    Author            = 'Simon Heather'

    # Company or vendor of this module
    CompanyName       = 'Guardian-Teck Ltd'

    # Copyright statement for this module
    Copyright         = 'Copyright 2020 Guardian-Teck ltd. All rights reserved.'

    # Description of the functionality provided by this module
    Description       = 'PowerShell Wrapper for the SSL Labs Assessment API'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @()

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport   = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData       = @{

        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags         = @('SSLLabs')

            # A URL to the license for this module.
            LicenseUri   = 'https://github.com/X-Guardian/SSLLabsScanPS/blob/master/LICENSE'

            # A URL to the main website for this project.
            ProjectUri   = 'https://github.com/X-Guardian/SSLLabsScanPS'

            # A URL to an icon representing this module.
            # IconUri      = ''

            # ReleaseNotes of this module
            ReleaseNotes = ''

            Prerelease   = ''
        } # End of PSData hashtable

    } # End of PrivateData hashtable
}
