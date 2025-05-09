Param(
    [parameter(Mandatory=${true})][string]${filePath},
	[parameter(Mandatory=${true})][string]${ucs}
)

process
{  
    Try
    {
        ${Error}.Clear()

        Import-Module "CiscoUcsPS"
        
        if(!(test-path -Path ${filePath}))
        {
                Write-Host "File Path : ${filePath} does not exist."
                exit
        }
          
        ${pathArr}= ${filePath}.Split("\")
        ${localFile}=  ${pathArr}[${pathArr}.Lengh-1]             
     
        ${handle} = Connect-Ucs ${ucs}
       
        ${uri} = [System.String]::Format("{0}/operations/file-{1}/image.txt", ${handle}.Uri, ${localFile})
           
        ${uploadFileOut} = [Cisco.Ucs.Utils]::UploadFile([Cisco.Ucs.UcsHandle]${handle},${uri}, ${localFile}, ${filePath}, ${null})

        if (${uploadFileOut} -ne ${null})
        {
            ${dn} = [Cisco.Ucs.Utils]::MakeDn([Cisco.Ucs.CapabilityCatalogue]::MakeRn(), [Cisco.Ucs.CapabilityEp]::MakeRn(), [Cisco.Ucs.CapabilityUpdater]::MakeRn(${localFile}))
            
            [Cisco.Ucs.ConfigMap] ${configMap} = new-object Cisco.Ucs.ConfigMap
            ${mo} = new-object Cisco.Ucs.CapabilityUpdater
            ${mo}.Dn = ${dn}
            ${mo}.Status = "created"
            ${mo}.FileName = ${localFile}
            ${mo}.Server="local"
            ${mo}.AdminState="restart"
            ${mo}.Protocol="local"

            ${pair} = new-object Cisco.Ucs.Pair
            ${pair}.Key = ${mo}.Dn
            ${pair}.AddChild(${mo})
            ${configMap}.AddChild(${pair})

            ${ccm} = [Cisco.Ucs.MethodFactory]::ConfigConfMos(${handle}, ${configMap}, "false", ${null})
            if(${ccm}.ErrorCode -ne 0)
            {
                write-host(${ccm}.ErrorDescr)
            }
			else
			{
				 Write-Host ""
				 Write-Host "Capability Catalog updated from a local file source."
			}
         }
       Disconnect-Ucs
        
    }
    Catch
    {
        Disconnect-Ucs
    	Write-Host ${Error}
       
    }
}