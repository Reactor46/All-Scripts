Param(
    [Parameter(Mandatory = ${true}, ValueFromPipeline = ${true})][Cisco.Ucs.ManagedObject[]] ${Mos}
)

process
{  
    Try
    {
    	${Error}.Clear()

    	foreach (${mo} in ${Mos})
    	{
    		${mo}.ToXml()
    	}
    }
    Catch
    {
    	Write-Host ${Error}
    }
}