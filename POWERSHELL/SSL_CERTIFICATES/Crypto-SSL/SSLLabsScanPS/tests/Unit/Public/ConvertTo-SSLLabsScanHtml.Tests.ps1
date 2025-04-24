<#
.Synopsis
   Tests for ConvertTo-SSLLabsScanHtml
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '',
    Justification = 'Suppress false positives in Pester code blocks')]
param()

$ProjectName = 'SSLLabsScanPS'

Import-Module $ProjectName

InModuleScope $ProjectName {
    Describe 'SSLLabsScanPS/ConvertTo-SSLLabsScanHtml' {
        Context 'When successfully invoked' {
            BeforeAll {
                $testEndPointData = Get-Content -Path "$PSScriptRoot\..\Resource\testEndpoint.json" |
                    ConvertFrom-Json
                $testEndPointData.details.hostStartTime = [DateTime]$testEndPointData.details.hostStartTime

                Mock -CommandName Out-File -RemoveParameterType Encoding

                $results = ConvertTo-SSLLabsScanHtml -EndPointData $testEndPointData
            }

            It 'Should have called the expected mocks' {
                Assert-MockCalled -CommandName Out-File -Exactly -Times 1
            }
        }
    }
}
