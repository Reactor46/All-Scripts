$script:apiProperties = @(
    @{
        ApiName  = 'Info'
        TypeName = 'SSLLabsScan.Info'
    }
    @{
        ApiName  = 'Analyze'
        TypeName = 'SSLLabsScan.HostData'
    }
    @{
        ApiName  = 'GetEndpointData'
        TypeName = 'SSLLabsScan.EndPointData'
    }
)

$script:baseEndpoint = 'https://api.ssllabs.com/api/v3/'

$script:resourceDirectoryName = $ExecutionContext.SessionState.Module.ModuleBase + '\Resources'
