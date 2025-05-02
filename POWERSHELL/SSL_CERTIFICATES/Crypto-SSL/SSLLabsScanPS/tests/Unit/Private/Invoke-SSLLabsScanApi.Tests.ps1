<#
.Synopsis
   Tests for Invoke-SSLLabsScanApi
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '',
    Justification = 'Suppress false positives in Pester code blocks')]
param()

$ProjectName = 'SSLLabsScanPS'

Import-Module $ProjectName

InModuleScope $ProjectName {
    Describe 'SSLLabsScanPS/Invoke-SSLLabsScanApi' {
        Context 'When successfully invoked' {
            BeforeAll {
                $testApiName = 'Info'

                Mock -CommandName Invoke-RestMethod -MockWith {
                    [PSCustomObject]@{
                        host = '1'
                    }
                }

                $results = Invoke-SSLLabsScanApi -ApiName $testApiName
            }

            It 'Should return the correct results' {
                $results | Should -BeOfType PSCustomObject
            }

            It 'Should have called the expected mocks' {
                Assert-MockCalled -CommandName Invoke-RestMethod -Exactly -Times 1
            }
        }
    }
}
