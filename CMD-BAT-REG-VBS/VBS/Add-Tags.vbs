' Copyright (c) 2012 Microsoft Corporation. All rights reserved.

' THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED "AS IS"
' WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' THE ENTIRE RISK OF USE, INABILITY TO USE, OR  RESULTS FROM
' THE USE OF THIS CODE REMAINS WITH THE USER.

Option Explicit

Dim errObjectNotCreated
errObjectNotCreated = 429

Sub GetActiveDirectoryInformation(ByRef strTag1, ByRef strTag2, ByRef strTag3)
    Dim objSystem
    Set objSystem = CreateObject("ADSystemInfo")
    If IsEmpty(objSystem) Then
        WScript.Quit(Err.Number)
    End If
    Dim objUser
    On Error Resume Next
    Set objUser = GetObject("LDAP://" & objSystem.UserName)
    On Error GoTo 0    
    If IsEmpty(objUser) Then
        WScript.Quit(errObjectNotCreated)
    ElseIf Err.Number <> 0 Then
        WScript.Quit(Err.Number)
    Else
        strTag1 = objUser.department                  ' Department Name
        strTag2 = objUser.title                       ' Job Title
        strTag3 = objUser.physicalDeliveryOfficeName  ' Office Name

        ' There are other common attributes, such as..
        ' objUser.c                                   ' Country/Region
        ' objUser.company                             ' Company
    End If
End Sub

Function CurrentUserDomain()
    Dim objWScriptNetwork
    Set objWScriptNetwork = CreateObject("WScript.Network")
    If IsEmpty(objWScriptNetwork) Then
        WScript.Quit(errObjectNotCreated)
    End If
    CurrentUserDomain = objWScriptNetwork.UserDomain
End Function

Sub WriteRegistry(ByRef objShell, ByVal strValueName, ByVal strValueData)
    Dim strPrefix

    strPrefix = "HKCU\Software\Microsoft\Office\15.0\osm\"
    objShell.RegWrite strPrefix & strValueName, strValueData, "REG_SZ"
End Sub

Function WScriptShell()
    Set WScriptShell = CreateObject("WScript.Shell")
    If IsEmpty(WScriptShell) Then
        WScript.Quit(errObjectNotCreated)
    End If
End Function

'
' Main script
'

Dim objShell
Set objShell = WScriptShell()

Dim strTag1, strTag2, strTag3
Call GetActiveDirectoryInformation(strTag1, strTag2, strTag3)

Dim strTag4
strTag4 = CurrentUserDomain()

Call WriteRegistry(objShell, "Tag1", strTag1)
Call WriteRegistry(objShell, "Tag2", strTag2)
Call WriteRegistry(objShell, "Tag3", strTag3)
Call WriteRegistry(objShell, "Tag4", strTag4)

'' SIG '' Begin signature block
'' SIG '' MIIaWAYJKoZIhvcNAQcCoIIaSTCCGkUCAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFMLThqoP/XGY
'' SIG '' sjpxmisgQKiIFBZVoIIVJjCCBJkwggOBoAMCAQICEzMA
'' SIG '' AACdHo0nrrjz2DgAAQAAAJ0wDQYJKoZIhvcNAQEFBQAw
'' SIG '' eTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EwHhcNMTIwOTA0
'' SIG '' MjE0MjA5WhcNMTMwMzA0MjE0MjA5WjCBgzELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjENMAsGA1UECxMETU9QUjEeMBwGA1UE
'' SIG '' AxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuqRJbBD7Ipxl
'' SIG '' ohaYO8thYvp0Ka2NBhnScVgZil5XDWlibjagTv0ieeAd
'' SIG '' xxphjvr8oxElFsjAWCwxioiuMh6I238+dFf3haQ2U8pB
'' SIG '' 72m4aZ5tVutu5LImTXPRZHG0H9ZhhIgAIe9oWINbSY+0
'' SIG '' 39M11svZMJ9T/HprmoQrtyFndNT2eLZhh5iUfCrPZ+kZ
'' SIG '' vtm6Y+08Tj59Auvzf6/PD7eBfvT76PeRSLuPPYzIB5Mc
'' SIG '' 87115PxjICmfOfNBVDgeVGRAtISqN67zAIziDfqhsg8i
'' SIG '' taeprtYXuTDwAiMgEPprWQ/grZ+eYIGTA0wNm2IZs7uW
'' SIG '' vJFapniGdptszUzsErU4RwIDAQABo4IBDTCCAQkwEwYD
'' SIG '' VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFN5R3Bvy
'' SIG '' HkoFPxIcwbzDs2UskQWYMB8GA1UdIwQYMBaAFMsR6MrS
'' SIG '' tBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeG
'' SIG '' RWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3Js
'' SIG '' L3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8wOC0zMS0yMDEw
'' SIG '' LmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKG
'' SIG '' Pmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2Vy
'' SIG '' dHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAuY3J0MA0G
'' SIG '' CSqGSIb3DQEBBQUAA4IBAQAqpPfuwMMmeoNiGnicW8X9
'' SIG '' 7BXEp3gT0RdTKAsMAEI/OA+J3GQZhDV/SLnP63qJoc1P
'' SIG '' qeC77UcQ/hfah4kQ0UwVoPAR/9qWz2TPgf0zp8N4k+R8
'' SIG '' 1W2HcdYcYeLMTmS3cz/5eyc09lI/R0PADoFwU8GWAaJL
'' SIG '' u78qA3d7bvvQRooXKDGlBeMWirjxSmkVXTP533+UPEdF
'' SIG '' Ha7Ki8f3iB7q/pEMn08HCe0mkm6zlBkB+F+B567aiY9/
'' SIG '' Wl6EX7W+fEblR6/+WCuRf4fcRh9RlczDYqG1x1/ryWlc
'' SIG '' cZGpjVYgLDpOk/2bBo+tivhofju6eUKTOUn10F7scI1C
'' SIG '' dcWCVZAbtVVhMIIEujCCA6KgAwIBAgIKYQKOQgAAAAAA
'' SIG '' HzANBgkqhkiG9w0BAQUFADB3MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EwHhcNMTIwMTA5MjIyNTU4WhcNMTMwNDA5MjIy
'' SIG '' NTU4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
'' SIG '' c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
'' SIG '' BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjENMAsGA1UE
'' SIG '' CxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
'' SIG '' OkY1MjgtMzc3Ny04QTc2MSUwIwYDVQQDExxNaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIBIjANBgkqhkiG
'' SIG '' 9w0BAQEFAAOCAQ8AMIIBCgKCAQEAluyOR01UwlyVgNdO
'' SIG '' Cz2/l0PDS+NgZxEvAU0M2NFGLxBA3gukUFISiAtDei0/
'' SIG '' 7khuZseR5gPKbux5qWojm81ins1qpD/no0P/YkehtLpE
'' SIG '' +t9AwYVUfuigpyxDI5tSHzI19P6aVp+NY3d7MJ4KM4Vy
'' SIG '' G8pKyMwlzdtdES7HsIzxj0NIRwW1eiAL5fPvwbr0s9jN
'' SIG '' OI/7Iao9Cm2FF9DK54YDwDODtSXEzFqcxMPaYiVNUyUU
'' SIG '' YY/7G+Ds90fGgEXmNVMjNnfKsN2YKznAdTUP3YFMIT12
'' SIG '' MMWysGVzKUgn2MLSsIRHu3i61XQD3tdLGfdT3njahvdh
'' SIG '' iCYztEfGoFSIFSssdQIDAQABo4IBCTCCAQUwHQYDVR0O
'' SIG '' BBYEFC/oRsho025PsiDQ3olO8UfuSMHyMB8GA1UdIwQY
'' SIG '' MBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRN
'' SIG '' MEswSaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNv
'' SIG '' bS9wa2kvY3JsL3Byb2R1Y3RzL01pY3Jvc29mdFRpbWVT
'' SIG '' dGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
'' SIG '' AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
'' SIG '' cGtpL2NlcnRzL01pY3Jvc29mdFRpbWVTdGFtcFBDQS5j
'' SIG '' cnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcN
'' SIG '' AQEFBQADggEBAHP/fS6dzY2IK3x9414VceloYvAItkNW
'' SIG '' xFxKLWjY+UgRkfMRnIXsEtRUoHWpOKFZf3XuxvU02FSk
'' SIG '' 4tDMfJerk3UwlwcdBFMsNn9/8UAeDJuA4hIKIDoxwAd1
'' SIG '' Z+D6NJzsiPtXHOVYYiCQRS9dRanIjrN8cm0QJ8VL2G+i
'' SIG '' qBKzbTUjZ/os2yUtuV2xHgXnQyg+nAV2d/El3gVHGW3e
'' SIG '' SYWh2kpLCEYhNah1Nky3swiq37cr2b4qav3fNRfMPwzH
'' SIG '' 3QbPTpQkYyALLiSuX0NEEnpc3TfbpEWzkToSV33jR8Zm
'' SIG '' 08+cRlb0TAex4Ayq1fbVPKLgtdT4HH4EVRBrGPSRzVGn
'' SIG '' lWUwggW8MIIDpKADAgECAgphMyYaAAAAAAAxMA0GCSqG
'' SIG '' SIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
'' SIG '' GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNV
'' SIG '' BAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1
'' SIG '' dGhvcml0eTAeFw0xMDA4MzEyMjE5MzJaFw0yMDA4MzEy
'' SIG '' MjI5MzJaMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
'' SIG '' YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
'' SIG '' VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNV
'' SIG '' BAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBMIIB
'' SIG '' IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAsnJZ
'' SIG '' XBkwZL8dmmAgIEKZdlNsPhvWb8zL8epr/pcWEODfOnSD
'' SIG '' GrcvoDLs/97CQk4j1XIA2zVXConKriBJ9PBorE1LjaW9
'' SIG '' eUtxm0cH2v0l3511iM+qc0R/14Hb873yNqTJXEXcr609
'' SIG '' 4CholxqnpXJzVvEXlOT9NZRyoNZ2Xx53RYOFOBbQc1sF
'' SIG '' umdSjaWyaS/aGQv+knQp4nYvVN0UMFn40o1i/cvJX0Yx
'' SIG '' ULknE+RAMM9yKRAoIsc3Tj2gMj2QzaE4BoVcTlaCKCoF
'' SIG '' MrdL109j59ItYvFFPeesCAD2RqGe0VuMJlPoeqpK8kbP
'' SIG '' Nzw4nrR3XKUXno3LEY9WPMGsCV8D0wIDAQABo4IBXjCC
'' SIG '' AVowDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQUyxHo
'' SIG '' ytK0FlgByTcuMxYWuUyaCh8wCwYDVR0PBAQDAgGGMBIG
'' SIG '' CSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYE
'' SIG '' FP3RMU7TJoqV4ZhgO6gxb6Y8vNgtMBkGCSsGAQQBgjcU
'' SIG '' AgQMHgoAUwB1AGIAQwBBMB8GA1UdIwQYMBaAFA6sgmBA
'' SIG '' VieX5SUT/CrhClOVWeSkMFAGA1UdHwRJMEcwRaBDoEGG
'' SIG '' P2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3Js
'' SIG '' L3Byb2R1Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBU
'' SIG '' BggrBgEFBQcBAQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6
'' SIG '' Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWlj
'' SIG '' cm9zb2Z0Um9vdENlcnQuY3J0MA0GCSqGSIb3DQEBBQUA
'' SIG '' A4ICAQBZOT5/Jkav629AsTK1ausOL26oSffrX3XtTDst
'' SIG '' 10OtC/7L6S0xoyPMfFCYgCFdrD0vTLqiqFac43C7uLT4
'' SIG '' ebVJcvc+6kF/yuEMF2nLpZwgLfoLUMRWzS3jStK8cOeo
'' SIG '' DaIDpVbguIpLV/KVQpzx8+/u44YfNDy4VprwUyOFKqSC
'' SIG '' HJPilAcd8uJO+IyhyugTpZFOyBvSj3KVKnFtmxr4HPBT
'' SIG '' 1mfMIv9cHc2ijL0nsnljVkSiUc356aNYVt2bAkVEL1/0
'' SIG '' 2q7UgjJu/KSVE+Traeepoiy+yCsQDmWOmdv1ovoSJgll
'' SIG '' OJTxeh9Ku9HhVujQeJYYXMk1Fl/dkx1Jji2+rTREHO4Q
'' SIG '' FRoAXd01WyHOmMcJ7oUOjE9tDhNOPXwpSJxy0fNsysHs
'' SIG '' cKNXkld9lI2gG0gDWvfPo2cKdKU27S0vF8jmcjcS9G+x
'' SIG '' PGeC+VKyjTMWZR4Oit0Q3mT0b85G1NMX6XnEBLTT+yzf
'' SIG '' H4qerAr7EydAreT54al/RrsHYEdlYEBOsELsTu2zdnnY
'' SIG '' CjQJbRyAMR/iDlTd5aH75UcQrWSY/1AWLny/BSF64pVB
'' SIG '' J2nDk4+VyY3YmyGuDVyc8KKuhmiDDGotu3ZrAB2WrfIW
'' SIG '' e/YWgyS5iM9qqEcxL5rc43E91wB+YkfRzojJuBj6DnKN
'' SIG '' waM9rwJAav9pm5biEKgQtDdQCNbDPTCCBgcwggPvoAMC
'' SIG '' AQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAwXzET
'' SIG '' MBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixk
'' SIG '' ARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0
'' SIG '' IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTA3
'' SIG '' MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkG
'' SIG '' A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
'' SIG '' BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
'' SIG '' dCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0
'' SIG '' IFRpbWUtU3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEF
'' SIG '' AAOCAQ8AMIIBCgKCAQEAn6Fssd/bSJIqfGsuGeG94uPF
'' SIG '' mVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKu
'' SIG '' mFwwJtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8
'' SIG '' e+eIHernTqOhwSNTyo36Rc8J0F6v0LBCBKL5pmyTZ9co
'' SIG '' 3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
'' SIG '' 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TT
'' SIG '' BoFKovmEpxFHFAmCn4TtVXj+AZodUAiFABAwRu233iNG
'' SIG '' u8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHOIktU
'' SIG '' //kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUw
'' SIG '' AwEB/zAdBgNVHQ4EFgQUIzT42VJGcArtQPt2+7MrsMM1
'' SIG '' sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcVAQQDAgEA
'' SIG '' MIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOV
'' SIG '' WeSkoWOkYTBfMRMwEQYKCZImiZPyLGQBGRYDY29tMRkw
'' SIG '' FwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
'' SIG '' EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRo
'' SIG '' b3JpdHmCEHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkw
'' SIG '' RzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
'' SIG '' L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNl
'' SIG '' cnQuY3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcw
'' SIG '' AoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
'' SIG '' ZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
'' SIG '' BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIB
'' SIG '' ABCXisNcA0Q23em0rXfbznlRTQGxLnRxW20ME6vOvnuP
'' SIG '' uC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0Bhqy7
'' SIG '' ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU
'' SIG '' 8gYvhQPpkSokInD79vzkeJkuDfcH4nC8GE6djmsKcpW4
'' SIG '' oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYxPStyC8jq
'' SIG '' cD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RT
'' SIG '' ciHuMMq7eGVcWwEXChQO0toUmPU8uWZYsy0v5/mFhsxR
'' SIG '' VuidcJRsrDlM1PZ5v6oYemIp76KbKTQGdxpiyT0ebR+C
'' SIG '' 8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpG
'' SIG '' QxS8s7OwTWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr
'' SIG '' 6bv0SMws1NgygEwmKkgkX1rqVu+m3pmdyjpvvYEndAYR
'' SIG '' 7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy
'' SIG '' /t9RyIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXU
'' SIG '' QS5QrJSwpXirs6CWdRrZkocTdSIvMqgIbqBbjCW/oO+E
'' SIG '' yiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
'' SIG '' XNhNG6HY+i+ePy5VFmvJE6P9MYIEnjCCBJoCAQEwgZAw
'' SIG '' eTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0ECEzMAAACdHo0n
'' SIG '' rrjz2DgAAQAAAJ0wCQYFKw4DAhoFAKCBwDAZBgkqhkiG
'' SIG '' 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgEL
'' SIG '' MQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU
'' SIG '' 6gcjgjUy20oZKUdJ4GR0pPNFBiowYAYKKwYBBAGCNwIB
'' SIG '' DDFSMFCgEoAQAEEAZABkAC0AVABhAGcAc6E6gDhodHRw
'' SIG '' Oi8vdGVjaG5ldC5taWNyb3NvZnQuY29tL2VuLXVzL2xp
'' SIG '' YnJhcnkvZWU4MTkwOTYuYXNweDANBgkqhkiG9w0BAQEF
'' SIG '' AASCAQBC9e3ko36WlhKi+kkjkYm1LT9h/zNDLav5UTnL
'' SIG '' deA9YAVf73DNMEZ2lxKC1n6WDSXGVCuNt+/tF2jZOfSX
'' SIG '' GqR/6RgExBqmoDbB+kDc9v1p3aGcWFPsQRU0ermb6Cbt
'' SIG '' fhWqqKItWbxr9q5mjPvrVkiRlH988xUxMzIbNBtF8RPj
'' SIG '' m7nIIC4nWzEuNas6fsicU4Ee85E5FIUbeKR1C9oPfrFc
'' SIG '' PO5ETiBLqAbsMKZjOM8wAVmOWZEWzvr4VbR9BqaneXgz
'' SIG '' dMy2QKtDSEEIkITX35NpydwCQwlGFN/Y34XZDq672Czc
'' SIG '' 1QsBex7VQCIaCxf0xwuJDndywpr2VwrzxiaWztMvoYIC
'' SIG '' HzCCAhsGCSqGSIb3DQEJBjGCAgwwggIIAgEBMIGFMHcx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQQIKYQKOQgAAAAAAHzAJ
'' SIG '' BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
'' SIG '' DQEHATAcBgkqhkiG9w0BCQUxDxcNMTIxMTEzMDUxNjMx
'' SIG '' WjAjBgkqhkiG9w0BCQQxFgQURduNkniVeAV4DJfF1Hpk
'' SIG '' PvucFWcwDQYJKoZIhvcNAQEFBQAEggEAUG6Q4AuI81po
'' SIG '' BcutwH83F6arRFEDkiUz9RUNNt3TQbjhJUop2Tk0lVP1
'' SIG '' tHbyuCGrXiSlffjNZJXskeyLFzOMc1TV9QKlKvmy4gop
'' SIG '' r6wK5xMefsV5CBsyg0Nw4q00bU430MewFsGmuoipCRxv
'' SIG '' P85HQukDku7wq7ERTcZssMniRsysCb4oQszqq9mZ0r/w
'' SIG '' sAj/qMRsqXQLsV2tU85sEyjwi35U6+2PFfJ44JzZxy4N
'' SIG '' zZ/57+Y2JoQi2nUU/hoJ2PHOrz7rk24givtnERJLtdAO
'' SIG '' U6JNPtnAsmjT8F8tsKYm/QfQyGF6MyIb95PSOvbJLk2Z
'' SIG '' xusVJU9hApskyFHBgFzxxQ==
'' SIG '' End signature block
