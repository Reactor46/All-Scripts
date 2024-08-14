param(
    [parameter(Mandatory=${true}, ValueFromPipeline = ${true})]
    [AllowEmptyString()]
    [string]${InputXml}
)
  
process
{
    Try
    {
    	${Error}.Clear()
        
              
        if (!([string]::IsNullOrEmpty([string]${InputXml}.Trim())))
        {         
    		${stringReader} = New-Object "System.Io.StringReader"(${InputXml})
    		${xmlReader} = [System.Xml.XmlReader]::Create(${stringReader})
    		${temp} = ${xmlReader}.Read()
    		${newObj} = New-Object ("Cisco.Ucs." + ${xmlReader}.Name)
    		${temp} = ${newObj}.LoadFromXml(${xmlReader})
    	    return ${newObj}
        }
    }
    Catch
    {
    	Write-Host ${Error}
    }
}