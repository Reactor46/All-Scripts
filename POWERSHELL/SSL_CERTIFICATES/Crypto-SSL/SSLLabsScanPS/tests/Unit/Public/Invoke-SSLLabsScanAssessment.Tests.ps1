<#
.Synopsis
   Tests for Invoke-SSLLabsScanAssessment
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '',
    Justification = 'Suppress false positives in Pester code blocks')]
param()

$ProjectName = 'SSLLabsScanPS'

Import-Module $ProjectName

InModuleScope $ProjectName {
    Describe 'SSLLabsScanPS/Invoke-SSLLabsScanAssessment' {
        Context 'When successfully invoked' {
            BeforeAll {
                $testHostName = 'www.bbc.co.uk'
                $mockHost = [PSCustomObject]@{
                    host      = $testHostName
                    startTime = '1593710865150'
                    testTime  = '1593711188979'
                    status    = 'Ready'

                }

                Mock -CommandName Invoke-SSLLabsScanApi -MockWith {
                    $mockHost
                }

                $result = Invoke-SSLLabsScanAssessment -HostName $testHostName
            }

            It 'Should return the correct results' {
                $result.host | Should -Be $testHostName
                $result.startTime | Should -BeOfType System.DateTime
                $result.testTime | Should -BeOfType System.DateTime
            }

            It 'Should have called the expected mocks' {
                Assert-MockCalled -CommandName Invoke-SSLLabsScanApi -Exactly -Times 1
            }
        }
    }
}
