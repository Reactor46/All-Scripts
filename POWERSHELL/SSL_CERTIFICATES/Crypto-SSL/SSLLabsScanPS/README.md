# SSLLabsScanPS

[![Build Status](https://dev.azure.com/simonheather99/SSLLabsScanPS/_apis/build/status/X-Guardian.SSLLabsScanPS?branchName=master)](https://dev.azure.com/simonheather99/SSLLabsScanPS/_build/latest?definitionId=8&branchName=master)
![Azure DevOps coverage (branch)](https://img.shields.io/azure-devops/coverage/simonheather99/SSLLabsScanPS/8/master)
[![Azure DevOps tests](https://img.shields.io/azure-devops/tests/simonheather99/SSLLabsScanPS/8/master)](https://simonheather99.visualstudio.com/SSLLabsScanPS/_test/analytics?definitionId=8&contextType=build)
[![PowerShell Gallery (with prereleases)](https://img.shields.io/powershellgallery/vpre/SSLLabsScanPS?label=SSLLabsScanPS%20Preview)](https://www.powershellgallery.com/packages/SSLLabsScanPS/)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/SSLLabsScanPS?label=SSLLabsScanPS)](https://www.powershellgallery.com/packages/SSLLabsScanPS/)

PowerShell Wrapper for the [SSL Labs Assessment API](https://github.com/ssllabs/ssllabs-scan/blob/stable/ssllabs-api-docs.md)

## Releases

For each merge to the branch `master` a beta release will be
deployed to [PowerShell Gallery](https://www.powershellgallery.com/).
Periodically a release version tag will be pushed which will deploy a
full release to [PowerShell Gallery](https://www.powershellgallery.com/).

## Installation

You can get the latest release of the SSLLabsScanPS module on the [PowerShell Gallery](https://www.powershellgallery.com/packages/SSLLabsScanPS)

```PowerShell
Install-Module -Name SSLLabsScanPS
```

## Usage

### Invoke a new assessment of a website and produce a summary HTML report

```PowerShell
$hostData = Invoke-SSLLabsScanAssessment -HostName www.mwam.com -StartNew -All Done
ConvertTo-SSLLabsScanHtml -EndpointData $hostData.endpoints
```

### Retrieve a cached assessment of a website and produce a summary HTML report

```PowerShell
$hostData = Invoke-SSLLabsScanAssessment -HostName www.mwam.com -FromCache -All Done
ConvertTo-SSLLabsScanHtml -EndpointData $hostData.endpoints
```

### Retrieve a cached assessment of a particular host of a website and produce a summary HTML report

```PowerShell
$endPointData = Get-SSLLabsScanEndpointData -HostName www.mwam.com -IPAddress '18.132.32.101' -FromCache
ConvertTo-SSLLabsScanHtml -EndpointData $endPointData
```

The default location of the summary HTML report is the user's `Documents` folder,
with a file name of `<hostName>-SSLLabsScanReport-<yyyyMMdd-HHmmss>.html`

## Documentation

For a full list of functions in SSLLabsScanPS and examples on their use, check
out the [SSLLabsScanPS wiki](https://github.com/X-Guardian/SSLLabsScanPS/wiki).

## Change log

A full list of changes in each version can be found in the [change log](CHANGELOG.md).

## Features

- Module built with [Module Builder](https://github.com/PoshCode/ModuleBuilder) and [Sampler](https://github.com/gaelcolas/Sampler).
- [Pester](https://github.com/pester/Pester) unit tests.
- Azure DevOps [CI/CD Pipeline](https://dev.azure.com/simonheather99/SSLLabsScanPS/_build?definitionId=8&_a=summary).
- Azure DevOps Test results report.
- Azure DevOps Code Coverage report.

## Licensing

SSLLabsScanPS is licensed under the [MIT license](LICENSE).
