<#
.Synopsis
   Tests for Get-SSLLabsScanEndpointData
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '',
    Justification = 'Suppress false positives in Pester code blocks')]
param()

$ProjectName = 'SSLLabsScanPS'

Import-Module $ProjectName

InModuleScope $ProjectName {
    Describe 'SSLLabsScanPS/Get-SSLLabsScanEndpointData' {
        Context 'When successfully invoked' {
            BeforeAll {
                $testHostName = 'www.bbc.co.uk'
                $testIPAddress = '1.1.1.1'

                $mockEndpointDetails = [PSCustomObject]@{
                    hostStartTime = '1593710865150'
                }

                $mockEndpoint = [PSCustomObject]@{
                    ipAddress = $testIPAddress
                    details   = $mockEndpointDetails
                }

                Mock -CommandName Invoke-SSLLabsScanApi -MockWith {
                    $mockEndpoint
                }

                $result = Get-SSLLabsScanEndpointData  -HostName $testHostName -IPAddress $testIPAddress
            }

            It 'Should return the correct results' {
                $result.host | Should -Be $testHostName
                $result.ipAddress | Should -Be $testIPAddress
                $result.details.hostStartTime | Should -BeOfType System.DateTime
            }

            It 'Should have called the expected mocks' {
                Assert-MockCalled -CommandName Invoke-SSLLabsScanApi -Exactly -Times 1
            }
        }
    }
}
