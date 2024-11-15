﻿#requires -Version 2.0

<#
    .SYNOPSIS
    Enabling Modern Authentication for Exchange Online
	
    .DESCRIPTION
    Enabling Modern Authentication for Exchange Online (Office 365)
	
    .EXAMPLE
    PS C:\> .\Enable-ModernAuth-Exchange.ps1
	
    .NOTES
    Works fine with Office 2013 and Office 2016 on Windows. Tested with Office 2016 on the Mac.
    You must enable it on your computers (Windows and Mac) as well! It is disabled by default.

    .LINK
    https://blogs.technet.microsoft.com/canitpro/2015/09/11/step-by-step-setting-up-ad-fs-and-enabling-single-sign-on-to-office-365/
#>
[CmdletBinding()]
param ()

begin
{
	# The Exchange Online URL
	$ExoURL = 'https://outlook.office365.com/powershell-liveid/'
	
	# Same as above, but for the German Office 365 (MCD)
	#$ExoURL = 'https://outlook.office.de/powershell-liveid/'
	
	# The Exchange Online Authentication method
	$ExoAuth = 'Basic'
}

process
{
	# Get the Credeantials (Could also be imported if you have dem saved)
	$credentials = (Get-Credential)
	
	# Create the new session
	$paramNewPSSession = @{
		ConfigurationName = 'Microsoft.Exchange'
		ConnectionUri	   = $ExoURL
		Credential		   = $credentials
		Authentication	   = $ExoAuth
		AllowRedirection  = $true
	}
	$ExoSession = (New-PSSession @paramNewPSSession)
	
	# Start the Session by importing it to the PowerShell Session
	$null = (Import-PSSession -Session $ExoSession)
	
	# Enable Modern Authentication, use $false to disable it
	$null = (Set-OrganizationConfig -OAuth2ClientProfileEnabled $true)
}

end
{
	# Cleanup
	$ExoSession = $null
}

#region License
<#
    BSD 3-Clause License

    Copyright (c) 2018, enabling Technology <http://enatec.io>
    All rights reserved.

    Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

    1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

    2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

    3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

    By using the Software, you agree to the License, Terms and Conditions above!
#>
#endregion License

#region Hints
<#
    This is a third-party Software!

    The developer(s) of this Software is NOT sponsored by or affiliated with Microsoft Corp (MSFT) or any of its subsidiaries in any way

    The Software is not supported by Microsoft Corp (MSFT)!
#>
#endregion Hints