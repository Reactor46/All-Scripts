<#
.Synopsis
   Tests for Get-SSLLabsScanAssessment
#>

[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '',
    Justification = 'Suppress false positives in Pester code blocks')]
param()

$ProjectName = 'SSLLabsScanPS'

Import-Module $ProjectName

InModuleScope $ProjectName {
    Describe 'SSLLabsScanPS/Get-SSLLabsScanInfo' {
        Context 'When successfully invoked' {
            BeforeAll {
                $results = Get-SSLLabsScanInfo
            }

            It 'Should return the correct results' {
                $results.PSTypeNames[0] | Should -Be 'SSLLabsScan.Info'
            }
        }
    }
}
