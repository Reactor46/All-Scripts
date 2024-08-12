# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 
 
# 
# Synopsis: This script imports Edge config that has been previously exported using
#   the export-edgeConfig script. 
# 
# Usage: 
#   
#   Validate Phase: .\importedgeconfig
#                   -cloneConfigData "C:\cloneConfigData.xml" 
#                   -cloneConfigAnswer "C:\cloneConfigAnswer.xml" 
#                   -isImport $false
#
#   Import Phase: .\importedgeconfig
#                   -cloneConfigData "C:\cloneConfigData.xml" 
#                   -cloneConfigAnswer "C:\cloneConfigAnswer.xml" 
#                   -isImport $true
#                   -Key:"A9ABA4D2C21C4bc58B303EA47BBE3608" (32 byte string used for
#                    password encryption/decryption)
#


param($cloneConfigData, 
       $cloneConfigAnswer, 
       $isImport = $false,
       $key = "A9ABA4D2C21C4bc58B303EA47BBE3608")
 
$cloneTemplateDoc = new-object System.Xml.XmlDocument
$cloneDataDoc = new-object System.Xml.XmlDocument
$cloneAnswerDoc = new-object System.Xml.XmlDocument

#############################################################################
# Logs the information to the cloneLogFile.log
# isHost : If true Write the message to the host, inaddition to the log file.
# message : The message string to be logged. 
#############################################################################

function Write-Information([Boolean] $isHost, [String] $message)
{
    $logger.WriteLine();
    $logger.WriteLine($message)        
    
    # Write the error message to the host also.
    if($isHost -eq $true)
    {
        write-host $message 
    }
}

###########################################################################
# The function adds some useful information into the log file
# before the actual Export process starts.
###########################################################################
function Write-LogStartInformation()
{
    $startTime = ([System.DateTime]::Now).ToString()
    Write-Information $false "************ BEGIN IMPORT PROCESS **************"
    Write-Information $false $startTime
}

#############################################################
# Releases the file handles used by the Export Script.
#############################################################
function Export-ReleaseHandles()
{
    # Restore Error Action Preference
    $ErrorActionPreference = $errorActionSave    

    write-debug "Export-ReleaseHandles"
    if($logger -ne $null)
    {  
        Write-Information $false "************ END IMPORT PROCESS **************"
        $logger.flush()
        $logger.close()    
    }
}

######################################################
# Retrieves the Root setup registry entry.
# returns: return entry value of found else null
######################################################
function GetEdgeInstallPath()
{
    write-debug "GetEdgeInstallPath"
    # Get the root setup entires.
    $setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
    $setupEntries = get-itemproperty $setupRegistryPath
    if($setupEntries -eq $null)
    {
        return $null
    }

    # Try to get the Install Path.
    $installPath = $setupEntries.MsiInstallPath
    
    return $installPath
}

######################################################
# Retrieves the Setup Regsitry value of the given entry.
# setupProperty: Setup Setting to the retrieved
# returns: return entry value of found else null
######################################################
function GetEdgeScriptsPath([String] $scriptFileName)
{
    write-debug "GetEdgeScriptsPath"

    # Append the script path with file name.
    $scriptFilePath = $script:installPath + "Scripts\" + $scriptFileName
    return $scriptFilePath
}


##################################################################################
# Creates a MultiValued string compatible with the Monad Multi value parameters.
# componentAnswerNodeItem: Multivalued node with child items.
# returns multivalued sting. ex: '0.0.0.0:25', '0.0.0.0:216'
##################################################################################

function createMultivaluedMonadString([System.Xml.XmlNode]$componentAnswerNodeItem)
{
    write-debug ("CreateMultivaluedMonadString ")
    $multiValuedString = $null
    $individualMultivaluedNodes = $componentAnswerNodeItem.SelectNodes("Item")
    write-debug ("individualNodes = $individualMultivaluedNodes")
    if ($IndividualMultivaluedNodes) 
    {
        # for each multivalued child node found.
        foreach($individualMultivaluedNode in $individualMultivaluedNodes)
        {
            write-debug ("nodeValue=$nodeValue")
            $nodeValue = $individualMultivaluedNode.Get_InnerXml().ToString().Trim()
            if($nodeValue -ne "")
            {
                 if($multiValuedString -eq $null)
                 {
                     $multiValuedString = "'" + $nodeValue + "'"
                     continue
                 }
                 $multiValuedString = $multiValuedString + ", '" + $nodeValue + "'"
                 
            }    
        }
    }
    return $multiValuedString
}

#######################################################################
# Validates the input parameters.
# Returns True if validation succeeds else False.
#######################################################################
    
function ValidateInput()
{
    write-debug ("ValidateInput")
    $isValid = $true

    #check if the Configuration Data file exists.
    $fileExists = [System.IO.File]::Exists($cloneConfigData)
    if($fileExists -eq $false)
    {
       Write-Information $true "Cannot find Config Data file[$cloneConfigData]"
       $isValid = $false
    }

    # if phase is Import and Answer File is passed as parameter.
    # make sure the file exists.
    if(($isImport -eq $true) -and ($cloneConfigAnswer -ne $null))
    {
        if([System.IO.File]::Exists($cloneConfigAnswer) -eq $false)
        {
            Write-Information $true "Cannot find Config Data file[$cloneConfigAnswer]" 
            $isValid = $false  
        }
    }

    if(($isImport -eq $false) -and ($cloneConfigAnswer -eq $null))
    {
        Write-Information $true "Please specify an Answer File Path." 
        $isValid = $false   
    }
    return $isValid
}

#########################################################
# Validates the IP Address
# address: Address to be validated
#########################################################

function ValidateIPAddress([String] $address)
{
    write-debug ("ValidateIPAddress " + $address)
    [System.Net.IPAddress] $IpAddress = [System.Net.IPAddress]::Parse($address)
    if($IpAddress -eq $null)
    {
        write-debug ("null IP Address")
        return $false     
    }

    # IPv4 or IPv6
    if(($IpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) -and
      ($IpAddress.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetworkV6))
    {
        write-debug ("not IPV4 or IPV6")
        return $false
    }

    #IPANY, IPv6ANY, Loopback or a configured address
    if(($IpAddress.Equals([System.Net.IPAddress]::Any) -eq $true) -or
       ($IpAddress.Equals([System.Net.IPAddress]::IPv6Any) -eq $true) -or
       ([System.Net.IPAddress]::IsLoopBack($IpAddress) -eq $true))
    {
        return $true
    }
    else
    {
        $flag = $false
        $ipHostEntry = new-object System.Net.IPHostEntry
        $ipHostEntry = [System.Net.Dns]::GetHostEntry([String]::Empty)

        #Lookup the address in the hostlist
        for($i = 0; $i -le $ipHostEntry.AddressList.Count; $i++)
        {
            if($IpAddress.Equals($ipHostEntry.AddressList[$i]) -eq $true)
            {
                $flag = $true
            }    
        }

        # bindingIP doesn't match any local IP Address
        if($flag -eq $false)
        {
            write-debug ("bindingIP doesn't match any local IP Address")
            return $false
        }
    }

    trap
    {
        return $false
    }
    return $true
}


#######################################################################
# Validate the Bindings Multivalued parameter.
# Bindings: The string to be validated against a given Type.
# Returns NULL if all the entires succeded.
# Returns HashTable with validation entries for each item.
#######################################################################

function ValidateBindings([String] $bindings)
{
    write-debug ("ValidateBindings " + $bindings)
    $ValidBindingListHash = @{}
    $validBindingsFlag = $true

    # Example : '0.0.0.0:25', '214.0.0.0:25' #

    # No need to check for Empty BindingsList.
    $bindingsList = $bindings.Trim().Split(',')
      
    foreach ($bindingsElement in $bindingsList)
    {
        # Split the IP:Port and validate the IPAddress Part.

        $bindingsElement = $bindingsElement.Replace("'", "").Trim() 
        $indexOfLastColon = $bindingsElement.LastIndexOf(':')
        if($indexOfLastColon -eq -1)
        {
            $ValidBindingListHash[$bindingsElement] = $false
            $validBindingsFlag = $false
            continue
        }

        $a = ValidateIPAddress($bindingsElement.SubString(0, $indexOfLastColon))
        $ValidBindingListHash[$bindingsElement] = $a
        if ($a -eq $false)
        {
            $validBindingsFlag = $false
        }
    }

    if ($validBindingsFlag -eq $true)
    {
        return $null
    }
    return $ValidBindingListHash
}

#########################################################
# Validates the local Server FQDN
# address: Local Server FQDN to be validated
#########################################################

function ValidateFqdn([String] $fqdn)
{
    write-debug ("ValidateFqdn " + $fqdn)

    $localServerName = [System.Net.Dns]::GetHostName()
    $localIPHostEntry = [System.Net.Dns]::GetHostEntry($localServerName)
    $localServerFQDN = $localIPHostEntry.HostName
    
    if ($fqdn -ine $localServerFQDN)
    {
        write-debug ("mismatched local Server FQDN")
        return $false     
    }

    trap
    {
        return $false
    }
    return $true
}


#######################################################################
# Validate whether a given Value matches to a given Type.
# Value: The string to be validated against a given Type.
# Type: The Validation Type.
# Returns True if validation succeeds else False.
# Returns false if the Type is "None"
#######################################################################

function ValidateItem([String] $value, [String] $type)
{
    write-debug ("ValidateItem Value:$value Type:$type")
    #Validation blocks for each Type.
    #New blocks should be added for new Types.

    $validItem = $true
    
    if ($type -eq "DirectoryPath")
    {
        $validItem = [System.IO.Directory]::Exists($value.Trim())
        write-debug "validItem=$validItem"

        if($validItem -eq $false)
        {
            $dirInfo = [System.IO.Directory]::CreateDirectory($value.Trim())
            
            $validItem = ($dirInfo -ne $null)

            # remove any newly created directory
            [System.IO.Directory]::Delete($value.Trim())
            trap
            {
                write-debug ("trap: set validItem false")
                $validItem = $false
            }
        }
    }
    elseif ($type -eq "NullableDirectoryPath")
    {
        if ($value -ne "")
        {
            $validItem = [System.IO.Directory]::Exists($value.Trim())
            write-debug "validItem=$validItem"

            if($validItem -eq $false)
            {
                $dirInfo = [System.IO.Directory]::CreateDirectory($value.Trim())
                
                $validItem = ($dirInfo -ne $null)
    
                # remove any newly created directory
                [System.IO.Directory]::Delete($value.Trim())
                trap
                {
                    write-debug ("trap: set validItem false")
                    $validItem = $false
                }
            }
            }
    }
    
    elseif ($type -eq "FilePath")
    {
        $FileFolder = [System.IO.Path]::GetDirectoryName($value.Trim())
        $validItem = ValidateItem $FileFolder "DirectoryPath"
    }
    elseif($type -eq "IPAddress")
    {
        $validItem = ValidateIPAddress($value.Trim())
    }
    elseif($type -eq "Bindings")
    {
        # The return value is either NULL or 
        # an Hash table with atleast one elements.
        # Multivalued property, we need to know the validation result
        # of individual elements. 
        $validItem = ValidateBindings($value.Trim())
        if ($validItem -eq $null)
        {
            write-debug ("valid bindings")
            $validItem = $true
        }
        else
        {
            write-debug ("invalid bindings")
            return $validItem
        }
    }
    elseif($type -eq "FQDN")
    {
        # The return value is either NULL or 
        # an Hash table with atleast one elements.
        # Multivalued property, we need to know the validation result
        # of individual elements. 
        $validItem = ValidateFqdn($value.Trim())
        if ($validItem -eq $null)
        {
            write-debug ("valid local server FQDN")
            $validItem = $true
        }
        else
        {
            write-debug ("invalid local server FQDN")
            return $validItem
        }
    }
    
    else
    {
        # if you are here XML Template has validationType 
        # set to something that we are not supporting.
        Write-Information $true "ValidationType [$type] is not supported."
    }
    write-debug ("return valid = $validItem")
    return $validItem
}


###############################################################################
# Retrieve all the components of give component type from the data file. 
# itterate through them to see if the machine specific information can be validated 
# on the Target machine. if validation returns false add a keyed entry into the
# answer file.
# component: component type to retrieve from Data file.
# cloneDataDoc: XML Data file
# returns: String representation of the XML answer components.
###############################################################################

function ValidateAndAppendComponentConfig([System.Xml.XmlElement]$component, 
                                          [System.Xml.XmlDocument]$cloneDataDoc)
{
    write-debug ""
    write-debug ("ValidateAndAppendComponentConfig " + $component.ToString())
    $componentPath = "/ExportedEdgeConfiguration/" + $component.ToString()
    $componentDataNodes = $cloneDataDoc.SelectNodes($componentPath)
    
    # If data nodes are present in the Data File
    $appendString = ""
    if($componentDataNodes)
    {
        #For each data node of this component Type
        foreach($componentDataNode in $componentDataNodes)
        {
            $individualComponentString = ""
            $componentItems = $component.Get_ChildNodes()
            $componentDataNodeItems = $componentDataNode.Get_ChildNodes()

            # if component has a key(meaning not unique, we have to add key to the
            # to the answer file node.
            if($component.GetAttribute("Key") -ne "None")
            {
                $individualComponentString = $individualComponentString + 
                                "<" + $component.ToString() + 
                                " " + $component.GetAttribute("Key")  +
                                "='" + $componentDataNode.Get_Item($component.GetAttribute("Key")).Get_InnerXml() + 
                                "'"+  ">"
            }
            # if the component doen't have a Key(always will be a single node, on both source
            # and target server
            else
            {
                $individualComponentString = $individualComponentString + "<" + 
                                  $component.ToString() + ">"
            }

            write-debug ("IndividualComponentString $individualComponentString `n")
            write-debug ""

            # for each machine specific item in this Data Node

            $machineSpecificAnswerString = ""
            foreach($componentItem in $componentItems)
            {
                #Get the validation Type Key
                $validationType = $componentItem.GetAttribute("ValidationType")

                write-debug ("Validate " + $componentItem.Get_InnerText() + " element of type $validationType" )  

                #validate the machine specific item on the target server.
                $isValid = ValidateItem $componentDataNode.Get_Item($componentItem.Get_InnerText()).Get_InnerXml() $validationType

                #if the validation failed add entry in the answer file
                # for admin to enter it manually.
                if ($isValid -eq $false)
                {
                    Write-Information $true ("Validation Failed for " + $componentItem.Get_InnerText() +
                                " element of type " + $validationType)

                    $machineSpecificAnswerString = $machineSpecificAnswerString + 
                        "<!-- Validation failed -->"

                    $machineSpecificAnswerString = $machineSpecificAnswerString + 
                        "<" + $componentItem.Get_InnerText() + ">" +
                        $componentDataNode.Get_Item($componentItem.Get_InnerText()).Get_InnerXml() +
                        "</" + $componentItem.Get_InnerText() + ">"
                }
                # if the validation succeeded Add an entry to the console.
                elseif ($isValid -eq $true)
                {
                    Write-Information $true ("Validation succeeded for " + $componentItem.Get_InnerText() +
                                " element of type $validationType" )
                }
                else
                {
                    # if we are here we got a hash table with validation status of each item
                    # of a multivalued property.

                    Write-Information $true ("Validation Failed for " + $componentItem.Get_InnerText() +
                                " Multivalued element of type $validationType : Answer file entry created" )

                    $machineSpecificAnswerString = $machineSpecificAnswerString + 
                            "<" + $componentItem.Get_InnerText() + ">"
                   
                    foreach($hashKey in $isValid.Get_Keys())
                    {
                        if ($isValid[$hashKey] -eq $false)
                        {
                            $machineSpecificAnswerString = $machineSpecificAnswerString + 
                                "<!-- Validation failed -->"
                        }
                        else
                        {
                            $machineSpecificAnswerString = $machineSpecificAnswerString + 
                                "<!-- Validation succeeded -->"
                        }
                        $machineSpecificAnswerString = $machineSpecificAnswerString + 
                            "<Item>" +
                            $hashKey + 
                            "</Item>"
                    }

                    $machineSpecificAnswerString = $machineSpecificAnswerString + 
                            "</" + $componentItem.Get_InnerText() + ">"
                }
            }
   
            write-debug ("machineSpecificAnswerString $machineSpecificAnswerString")
            if ($machineSpecificAnswerString -ne "")
            {
                $appendString = $appendString + $individualComponentString
                $appendString = $appendString + $machineSpecificAnswerString
                $appendString = $appendString + "</" + $component.ToString() + ">"
            }
        }
    }
    return $appendString
}

##############################################################################
# This function does the extra processing work, which is required for some
# parameters outputted by the Monad Tasks.
# DataComponent: The data node from which we will populate
#                the values for the task parameters.
# Returns: The command string if any else null.
##############################################################################

function CreateExtraProcessingCommand([System.Xml.XmlElement]$datacomponent)
{
    $extraCommandString = $null
    
    # As of now SendConnector is the only object we have to do extra processing.
    # Only for the LinkedReceiveConnector parameter.

    if($datacomponent.ToString().ToUpper() -eq "SENDCONNECTOR")
    {
        write-debug ("CreateExtraProcessingCommand " + $dataComponent.ToString())
        $linkedReceiveConnector = $datacomponent.Get_Item("LinkedReceiveConnector").Get_InnerXml().Trim()
        if($linkedReceiveConnector -ne '$null')
        {
            $extraCommandString = "Set-SendConnector -Identity '" + 
                                  $datacomponent.Get_Item("Name").Get_InnerXml().Trim() +
                                  "'"
            $linkedReceiveConnectorItems = $linkedReceiveConnector.Trim().Split(',')
            $linkedReceiveConnectorObject = get-receiveconnector $linkedReceiveConnectorItems[0]
            if($linkedReceiveConnectorObject -ne $null)
            {
                $extraCommandString = $extraCommandString +
                                      " -LinkedReceiveConnector '" + 
                                      $linkedReceiveConnectorObject.Identity +
                                      "'"
            }
        }       
    }
    return $extraCommandString
}

##########################################################################################
# Create each individual monad task command to be run on the Target Server.
# command: The command name, its also used to get the parameters this command takes.
# Datacomponent: The data node from which we will populate the values for the task parameters.
# AnswerComponent: The answer node from which we will populate the machine related task paramaters.
# NonClonableItemsNode: Skip the task parameters which are listed in this Node.
# Key: There are some commands whose identity or server info cannot be cloned and they are single
#       Objects on Edge, in these cases use pipelining to set the Object
##########################################################################################

function EvaluateAndAppendCommandParameters([String]$command, 
                                            [System.Xml.XmlElement]$datacomponent, 
                                            [System.Xml.XmlElement]$answerComponent, 
                                            [System.Xml.XmlElement]$NonClonableItemsNode, 
                                            [String]$key,
                                            [String]$preConditionVerb)
                                            
{
    write-debug ("EvaluateAndAppendCommandParameters " + $command)
    write-debug ("preconditionVerb=" + $preConditionVerb)
   
    $commandWithParameters = ""
    $cmdObject = get-command $command
    $commandWithParameters = $command

    # for each parameter of this object
    foreach($parameter in $cmdObject.ParameterSets[0].Parameters)
    {
        write-debug ("parameter= " + $parameter.Name)

        # For each parameter the resulting command can take;
        # If the parameter is in the non-clonable list just ignore it
        # Else Try to get the value from the Answer file if any.
        # Else use the value present in the data file.
        if($NonClonableItemsNode)
        {
            $nonClonableItems = $NonClonableItemsNode.Item
            $nonCloneFlag = $false
            
            # if this parameter is a no clonnable item,
            # ignore this parameter in resulting task command.
            foreach($nonClonableItem in $nonClonableItems)
            {
                if($nonClonableItem.Trim() -eq $parameter.Name)
                {
                    $nonCloneFlag = $true
                    break;
                }
            }
        
            if($nonCloneFlag -eq $true)
            {
                continue;
            }
        }


        # if Data for this parameter is present in the Answer file pick it
        # else use the data from the Data file.
        # we don't need to check for "" and "{}" since the intial validation
        # while importing will fail if they are empty.
        if(($answerComponent) -and 
           ($answerComponent.Get_Item($parameter.Name)) -and 
           ($answerComponent.Get_Item($parameter.Name).Get_InnerXml() -ne "") -and 
           ($answerComponent.Get_Item($parameter.Name).Get_InnerXml() -ne "{}"))
        {
            # if Multivalued Answer Node create the command string from the
            # child nodes.
            if ($answerComponent.Get_Item($parameter.Name).Get_ChildNodes().Count -gt 1)
            {
                $multiValuedString = createMultivaluedMonadString($answerComponent.Get_Item($parameter.Name))
                $paramString = " -" + 
                       $parameter.Name + ":" + 
                       $multiValuedString    

                $commandWithParameters = $commandWithParameters + $paramString
            }
            # if boolean use $true or $false. Based on its value in the answer file
            elseif($parameter.ParameterType.Name -eq "Boolean")
            {
                $commandWithParameters = $commandWithParameters + " -" + 
                       $parameter.Name + ":$" + 
                       $answerComponent.Get_Item($parameter.Name).Get_InnerXml().ToLower()
            }
            # if not boolean consider as a regular string, tasks will take care of
            # converting them to specific Type ( TimeStamp, IpAddress etc..)
            else
            {
                $escapedValue = $answerComponent.Get_Item($parameter.Name).Get_InnerXml().Trim()
                $escapedValue = $escapedValue.Replace("&gt;",">")
                $escapedValue = $escapedValue.Replace("&lt;","<")

                $commandWithParameters = $commandWithParameters + " -" + 
                       $parameter.Name + " '" + 
                       $escapedValue + "'"
            }        
        }
        # if not present in the answer file, pick them from the data file only
        # if the parameter is there 
        elseif ($datacomponent.Get_Item($parameter.Name) -ne $null) 
        {
            write-debug ("parameter elt present= " + $parameter.Name)

            # if it is not blank OR the precondition is "SetBlank".
            if (($datacomponent.Get_Item($parameter.Name).Get_InnerXml() -ne "") -or
                ($preConditionVerb -eq "SetBlank"))
            {

                write-debug ("parameter not blank OR no Remove precondition")
          
                # pass null to the corresponding element during import phase.
                if($datacomponent.Get_Item($parameter.Name).Get_InnerXml() -eq '$null')
                {
                    $commandWithParameters = $commandWithParameters + " -" + 
                           $parameter.Name + ":$" + "null"            
                }
                # if the parameter is of type MultiValued construct the command 
                # accordingly.
                elseif($parameter.ParameterType.Name.ToString().ToUpper().StartsWith("MULTIVALUED"))
                {
                    $paramString = " -" + 
                           $parameter.Name + ":" + 
                           $datacomponent.Get_Item($parameter.Name).Get_InnerXml().Trim()

                    $commandWithParameters = $commandWithParameters + $paramString
                }
                # if boolean use $true or $false. Based on its value in the Data node.
                elseif($parameter.ParameterType.Name -eq "Boolean")
                {
                    $commandWithParameters = $commandWithParameters + " -" + 
                           $parameter.Name + ":$" + 
                           $datacomponent.Get_Item($parameter.Name).Get_InnerXml().ToLower()
                }
                # if PSCredential construct a new credential object
                elseif($parameter.ParameterType.Name -eq "PSCredential")
                {
                    $mshCredentialNode = $datacomponent.Get_Item($parameter.Name)

                    $userName = $mshCredentialNode.Get_Item("UserName").Get_InnerXml()	
                    $encryptedPassword = $mshCredentialNode.Get_Item("Password").Get_InnerXml()

                    $securePassword = ConvertTo-SecureString -String $encryptedPassword -key $encryptionKey
#               $encryptedPassword = ConvertFrom-SecureString -SecureString $securePassword -key $encryptionKey

                    # Catch password conversion errors
                    if ($securePassword -eq $null)
                    {
                        throw $Error[0].Exception
                    }

                    $script:Credential = new-object System.Management.Automation.PSCredential ($userName, $securePassword)

                    write-debug ("credential.UserName=" + $Credential.UserName)
#                write-debug ("credential.Password=" + $Credential.GetNetworkCredential().Password)

                    $paramString = " -" + 
                           $parameter.Name + ":" + '$script:Credential'

                    $commandWithParameters = $commandWithParameters + $paramString
                }
                elseif ($parameter.Name -eq "ProxyServerType")
                {
                    $stringProxyServerType = $datacomponent.Get_Item($parameter.Name).Get_InnerXml().ToLower()
                    
                    write-debug ("stringProxyServerType = " + $stringProxyServerType)                                

                    $paramString = " -" + 
                           $parameter.Name + ":" + $stringProxyServerType

                    $commandWithParameters = $commandWithParameters + $paramString
                }
                # Otherwise consider as a regular string, tasks will take care of
                # converting them to specific Type ( TimeStamp, IpAddress etc..)
                else
                {
                    $escapedValue = $datacomponent.Get_Item($parameter.Name).Get_InnerXml().Trim()
                    $escapedValue = $escapedValue.Replace("&gt;",">")
                    $escapedValue = $escapedValue.Replace("&lt;","<")

                    $commandWithParameters = $commandWithParameters + " -" + 
                           $parameter.Name + " '" + 
                           $escapedValue + "'"
                }
            }
        }
    }

    # If this command has machine specific information and the component
    # associated with this is unique. ex-- TransportServer.
    if(($key) -and ($key -eq "None"))
    {
        $MachineName = [System.Environment]::MachineName
        $commandWithParameters = "Get-" + $datacomponent.ToString() + " -Identity:" +
                                 $MachineName + " | " +
                                 $commandWithParameters
        #write-debug ("Unique/Machine specific completeCommand= " + $commandWithParameters)
    }

    Write-Information $false $commandWithParameters
    write-debug ("commandWithParameters= " + $commandWithParameters)
    return $commandWithParameters
}

##########################################################################################
# Creates and runs commands
# ImportTemplateComponent: Encapsulates the Type of component and 
# commands to be run on the Target machine.
# Datacomponent: Data File Node
# AnswerComponent: Machine Specific Data Node
# Key: Key which Distinguishes objects of same component Type.
##########################################################################################

function ImportObject([System.Xml.XmlElement]$importTemplateComponent, 
                       [System.Xml.XmlElement]$datacomponent, 
                       [System.Xml.XmlElement]$answerComponent, 
                       [String]$key)
{
    write-debug ""
    write-debug ("ImportObject " + $importTemplateComponent.ToString())
    $preConditionVerb = $importTemplateComponent.GetAttribute("PreCondition").ToString()
    write-debug ("Precondition = " + $preconditionVerb)
    
    # Creates the Monad Command
    # As of now the PreCondition and PostCOndition have only single command.
    # ImportVerbs may have multiple elements so split them and build each command seperately.

    $command = ""
    $baseCommand = ""

    if(($importTemplateComponent.GetAttribute("ImportVerbs").ToString()) -ne "None")
    {
        # In some cases output from a single task may have to be
        # sent to multiple tasks with the same Noun.
        # In these cases ImportVerbs will have multiple items separated by :
        $subCmds = $importTemplateComponent.GetAttribute("ImportVerbs").ToString().Split(':')

        write-debug ("ImportVerbs:" + $importTemplateComponent.GetAttribute("ImportVerbs").ToString())
        foreach($subCmd in $subCmds)
        {
            # Special Case specific to TransportAgents
            # There are two different tasks to enable/disable(why?).
            if($subCmd -eq "EnableOrDisable")
            {
                #Assuming that the parameter name associated with flag is "Enabled"
                #Will have to add some extra mappings in Template XML to care of this.
                if($datacomponent.Get_Item("Enabled").Get_InnerXml().ToString() -eq "True")
                {
                    $baseCommand = "Enable-" + $importTemplateComponent.ToString()
                }
                else
                {
                    $baseCommand = "Disable-" + $importTemplateComponent.ToString()
                }
            }
            else
            {
                # command created from ImportVerbs element and component name which is Noun.
                $baseCommand = $subCmd + "-" + $importTemplateComponent.ToString()
            }

            # Now we have command name, So go get the parameters and append to command.
            $nonclonableItemsNode = $importTemplateComponent.Get_Item("NonClonableItems") 

            $completeCommand = EvaluateAndAppendCommandParameters $baseCommand $datacomponent $answerComponent $nonClonableItemsNode $key $preConditionVerb

            write-debug ("importTemplateComponent = " + $importTemplateComponent.ToString())

            # Special Case specific to RemoteDomain to deal with Identity and unremovable Default instance
            if ($importTemplateComponent.ToString() -eq "RemoteDomain") 
            {
                $domainName = $datacomponent.Get_Item("DomainName").Get_InnerXml().ToString()
                write-debug "domainName=$domainName"                    

                #We can't do a New on the (fixed) Default RemoteDomain Value since it can't be removed initially
                if ($subCmd -eq "New") 
                {
                    if ($domainName -eq "*")
                    {
		        # Special handling for Default Remote Domain
                        # We can't do a New on the Default RemoteDomain Value since it can't be removed initially
			# However if it doesn't exist we have to create it	
                        $remoteDomain = get-remotedomain | where {$_.DomainName -eq $domainName} 
			if ($remoteDomain -ne $null)
			{
	                        write-debug "Continue for Default New-RemoteDomain"
	                        continue
			}
                    }
                }
                else
                {
                    # Save the New-RemoteDomain Identity param for the Set-RemoteDomain
                    $remoteDomain = get-remotedomain | where {$_.DomainName -eq $domainName} 
                    write-debug ("Identity=" + $remoteDomain.Identity.Name.ToString())

                    $completeCommand = $completeCommand + " -Identity:" + "'" + $remoteDomain.Identity.Name + "'" 
                }
            }

            # Special Case specific to AcceptedDomain to deal with Identity 
            if ($importTemplateComponent.ToString() -eq "AcceptedDomain") 
            {
                $domainName = $datacomponent.Get_Item("DomainName").Get_InnerXml().ToString()
                write-debug "domainName=$domainName"                    

                $name = $datacomponent.Get_Item("Name").Get_InnerXml().ToString()
                write-debug "Name=$name"                    

                #We can't do a New on the (fixed) Default RemoteDomain Value since it can't be removed initially
                if ($subCmd -eq "Set") 
                {
                    # Save the New-AcceptedDomain Identity param for the Set-AcceptedDomain
                    $acceptedDomain = get-accepteddomain | where {$_.Name -eq $name}
                    write-debug ("Identity=" + $acceptedDomain.Identity.Name.ToString())

                    $completeCommand = $completeCommand + " -Identity:" + "'" + $acceptedDomain.Identity.Name + "'" 

                    # If this was the Original Default Accepted Domain make it the new Default
                    write-debug ("defaultDomain?:" + $datacomponent.Get_Item("Default").Get_InnerXml()) 
                    if ($datacomponent.Get_Item("Default").Get_InnerXml() -eq $true) 
                    {
                        write-debug ("Original Default AcceptedDomain")
                        $completeCommand = $completeCommand + ' -MakeDefault:$true'
                    }
                }
            }

            write-debug ("invoke-expression " + $completeCommand)            

            # invoke the command with parameters.
            invoke-expression $completeCommand
            $command = $command + $baseCommand
        }

        # This is to take care of some parameters that cannot be handled 
        # in a generic way. EvaluateAndAppendCommandParameters only looks at the first parameter set
        # For SendConnectors the LinkedReceiveConnector parameter is in the second parameter set and so
        # this needs to be handled as a special case.
        
        $extraCommand = CreateExtraProcessingCommand $datacomponent
        # write-debug ("extraCommand= " + $extraCommand)
        if($extraCommand -ne $null)
        {
            Write-Information $false $extraCommand
            invoke-expression $extraCommand
            $command = $command + $extraCommand    
        }
    }

    return $command
}


##########################################################################################
# Imports the configuration using clone Template, Data File and Answer File.
# cloneTemplateDoc: Clone Template document.
# cloneDataDoc: Data document.
# cloneAnswerDoc: Answer file document.
##########################################################################################

function ImportConfig([System.Xml.XmlDocument]$cloneTemplateDoc, 
                      [System.Xml.XmlDocument]$cloneDataDoc, 
                      [System.Xml.XmlDocument]$cloneAnswerDoc)
{
    write-debug ("ImportConfig ")
    $importSpecificRoot = $cloneTemplateDoc.SelectSingleNode("/cloneTemplate/importCloneItems")
    # Get all the import items.
    $importItems = $importSpecificRoot.Get_ChildNodes()
    # Get all machine specific items.
    $machineSpecificRoot = $cloneTemplateDoc.SelectSingleNode("/cloneTemplate/machineSpecificItems")

    # For each import specific item.
    foreach($importItem in $importItems)
    {
        $componentDataPath = "/ExportedEdgeConfiguration/" + $importItem.ToString()
        $AnswerDataPath = "/MachineSpecificSettings/" + $importItem.ToString()

        # Get the machine specific node if any for this import item.
        $machineSpecificComponentItem = $machineSpecificRoot.Get_Item($importItem.ToString())

        # Get the Data nodes if any for this import item.
        $componentSpecificDataNodes = $cloneDataDoc.SelectNodes($componentDataPath)

        # Get the Answer file node if any for this import item.
        $componentSpecificAnswerNodes = $cloneAnswerDoc.SelectNodes($AnswerDataPath)

        write-debug ""
        write-debug ("importItem= " + $importItem.ToString())
        write-debug ("componentSpecificDataNodes= $componentSpecificDataNodes")
        write-debug ("componentSpecificAnswerNodes= $componentSpecificAnswerNodes")

        $key = $null
        if($machineSpecificComponentItem)
        {
            # Get the key if any for this import item from machine specfic Node.
            $key = $machineSpecificComponentItem.GetAttribute("Key")
        }

        $preConditionVerb = $importItem.GetAttribute("PreCondition").ToString()
        $componentNoun = $importItem.ToString()
        write-debug ("key= $key")
        write-debug ("PreconditionVerb = " + $preConditionVerb)

        #As of now precondition can be either "None", "SetBlank"  or "Remove"
        #If "Remove" Remove all the components of the current Type before adding new ones.
        #Better way of doing is put all the cleanup process as a separate block in the Template
        if( $preConditionVerb -eq "Remove")
        {
            # Clean up all the items of this Type on the Target machine.
            # I prefered to do it here so that we can delete all the objects in a single command
            # by using Monad pipelining
            if ($componentNoun -eq "IPAllowListEntry")
            {
                $baseCommand = $preConditionVerb + "-" + $componentNoun
                $cleanupCommand = 'Get-IPAllowListEntry | where { $_.IsMachineGenerated -eq $false}' + " | " + $baseCommand
                write-debug ("cleanupCommand= $cleanupCommand")
                invoke-expression $cleanupCommand  
            }
            elseif ($componentNoun -eq "RemoteDomain")
            {
                $baseCommand = $preConditionVerb + "-" + $componentNoun
                $cleanupCommand = 'Get-RemoteDomain | where { $_.DomainName -ne "*"}' + " | " + $baseCommand
                write-debug ("cleanupCommand= $cleanupCommand")
                invoke-expression $cleanupCommand  
            }
            elseif ($componentNoun -eq "AcceptedDomain")
            {
                # If no accepteddomains exist then don't add the unique one.
                $existingDomains = Get-AcceptedDomain
                
                if ($existingDomains -ne $null)
                {
                    # Add a Unique GUID Value and make it the new Default
                    $cleanupCommand = 'New-AcceptedDomain '
                    $cleanupCommand = $cleanupCommand + ' -DomainName:' + $script:uniqueName
                    $cleanupCommand = $cleanupCommand + ' -Name:' + $script:uniqueName
                    write-debug ("cleanupCommand= $cleanupCommand")
                    invoke-expression $cleanupCommand  

                    $cleanupCommand = 'Set-AcceptedDomain -Identity:' + $script:uniqueName + ' -MakeDefault:$true'
                    write-debug ("cleanupCommand= $cleanupCommand")
                    invoke-expression $cleanupCommand  

                    $baseCommand = $preConditionVerb + "-" + $componentNoun
                    $cleanupCommand = 'Get-AcceptedDomain | where { $_.Default -ne $true}' + " | " + $baseCommand
                    write-debug ("cleanupCommand= $cleanupCommand")
                    invoke-expression $cleanupCommand  
                }
            }
            else
            {
                $baseCommand = $preConditionVerb + "-" + $componentNoun
                $cleanupCommand = "Get-" + $componentNoun + " | " + $baseCommand
                write-debug ("cleanupCommand= $cleanupCommand")
                invoke-expression $cleanupCommand
            }
        }

        $cmd = ""
        # Get the data nodes for each import specific item. 
        # and import the data
        foreach($componentSpecificDataNode in $componentSpecificDataNodes)
        {
            [Boolean]$foundMatchingAnswerNode = $false
            # find a matching Answer node if any.
            foreach($componentSpecificAnswerNode in $componentSpecificAnswerNodes)
            {
                 # if the key is "None", unique node so just pass whatever you have in hand and break.
                 if($key -eq "None")
                 {
                    $foundMatchingAnswerNode = $true
                    write-debug ("found matching answer node for unique data node")
                    $cmd = ImportObject $importItem $componentSpecificDataNode $componentSpecificAnswerNode $key
                    break
                 }
                 
                 # find a matching key node in the answer file for the component and use that Answer Node.
                 elseif(($componentSpecificAnswerNode.GetAttribute($key)) -eq 
                        ($componentSpecificDataNode.Get_Item($key).Get_InnerXml()))
                 {
                     write-debug ("found matching answer node for multi data node")
                     $foundMatchingAnswerNode = $true
                     $cmd = ImportObject $importItem $componentSpecificDataNode $componentSpecificAnswerNode $key
                     break;    
                 }
            }

            # If there is no answer node(meaning doesn't contain any machine specific information
            # for this item) just use the Data Node.    
            if($foundMatchingAnswerNode -eq $false)
            {
                write-debug ("no machine specific answer node - use exported data node directly")
                $cmd = ImportObject $importItem $componentSpecificDataNode $null $key
            }
        }  

        #As of now postcondition can be either "None" or "Remove"
        #If "Remove" then Remove the temporary Unique Default value used to handle non-removeble defaults
        $postConditionVerb = $importItem.GetAttribute("PostCondition").ToString()
        $componentNoun = $importItem.ToString()
        write-debug ("PostConditionVerb = " + $postConditionVerb)
        if( $postConditionVerb -eq "Remove")
        {
            #As of now only AcceptedDomain has a postcondition 
            if ($componentNoun -eq "AcceptedDomain")
            {
                $postConditionCommand = 'Get-AcceptedDomain | where {$_.DomainName -eq "' + $script:uniqueName + '"} | Remove-AcceptedDomain '

                write-debug ("postConditionCommand= $postConditionCommand")
                invoke-expression $postConditionCommand  
            }
        }
    }    
}

##########################################################################################
# Creates the Answer File.
# cloneTemplateDoc: Clone Template document.
# cloneDataDoc: Data document.
# cloneConfigAnswer: Name of the answer file including complete path.
##########################################################################################

function ValidateAndCreateAnswerFile([System.Xml.XmlDocument]$cloneTemplateDoc, 
                                     [System.Xml.XmlDocument]$cloneDataDoc, 
                                     [String]$cloneConfigAnswer)
{
    write-debug "ValidateAndCreateAnswerFile "
    $rootNode = "MachineSpecificSettings"
    $xmlData = "<" + $rootNode + ">"
    $machineSpecificNode = $cloneTemplateDoc.SelectSingleNode("/cloneTemplate/machineSpecificItems")

    # Get all the machine specific nodes.
    $components = $machineSpecificNode.Get_ChildNodes()

    # for each machine specific node
    foreach($component in $components)
    {
        # Validate the machine specific info on the Target machine and
        # Get all the non valid nodes.
        $xmlItem = ValidateAndAppendComponentConfig $component $cloneDataDoc
        $xmlData = $xmlData + $xmlItem
    }

    # There should not be any string overflow because there is very little data 
    # in the Answer file.
    $xmlData = $xmlData + "</" + $rootNode + ">"
    $xmlDocument = [System.Xml.XmlDocument]($xmlData)
    $xmlDocument.Save($cloneConfigAnswer)
}




##########################################################################################
# Validates the Answer File.
# cloneTemplateDoc: Clone Template document.
# cloneDataDoc: Data document.
# Returns True if success else False
##########################################################################################

function ValidateAnswerFile([System.Xml.XmlDocument]$cloneTemplateDoc, 
                            [System.Xml.XmlDocument]$cloneAnswerDoc)
{
    write-debug "ValidateAnswerFile"
    write-verbose "ValidateAnswerFile"
    $isValid = $true
    $machineSpecificNode = $cloneTemplateDoc.SelectSingleNode("/cloneTemplate/machineSpecificItems")
    write-debug ("machineSpecificNode= " + $machineSpecificNode)

    # Get Machine specific componenets
    $components = $machineSpecificNode.Get_ChildNodes()

    # for each machine specific component.
    foreach($component in $components)
    {
        write-debug ("TemplateComponent = " + $component.ToString())
        $componentItems = $component.Get_ChildNodes()
        $componentPath = "/MachineSpecificSettings/" + $component.ToString()
        
        # Get all the components for this machine specific component.
        $componentAnswerNodes = $cloneAnswerDoc.SelectNodes($componentPath)
        if($componentAnswerNodes)
        {
            # For each answer component
            foreach($componentAnswerNode in $componentAnswerNodes)
            {
                write-debug ("AnswerFileComponent = " + $componentAnswerNode.ToString())

                # Get each individual item in this Answer component node
                $componentAnswerNodeItems = $componentAnswerNode.Get_ChildNodes()
                if($componentAnswerNodeItems)
                {
                    # Do process for each item in this answer node.
                    # for MultiValuedProperty node we will have child nodes
                    # which have to be taken in consideration.
                    foreach($componentAnswerNodeItem in $componentAnswerNodeItems)
                    {     
                        write-debug ("AnswerNodeItem=" + $componentAnswerNodeItem.ToString())

                        #Loop through the machine items to find a match
                        # for this answer item
                        foreach($componentItem in $componentItems)
                        {

                            write-debug ("TemplateNodeItem=" + $componentItem.Get_InnerXml().ToString())
                            # if match is found
                            if($componentItem.Get_InnerXml().ToString() -eq $componentAnswerNodeItem.ToString())
                            {
                                write-debug ("MATCH:TemplateNode=AnswerNode")
                                # if Multivalued property Validation
                                if ($componentAnswerNodeItem.Get_ChildNodes().Count -gt 1)
                                {
                                    $currentAnswerNodeValue = createMultivaluedMonadString($componentAnswerNodeItem)
                                }
                                else
                                {
                                    $currentAnswerNodeValue = $componentAnswerNodeItem.Get_InnerXml().ToString().Trim()
                                }
                                write-debug ("AnswerNodeValue=$currentAnswerNodeValue")

                                #if data not found for this item in the answer node validation failed.
                                if ($currentAnswerNodeValue -eq "")
                                {
                                    Write-Information $true "Validation Failed: " $component.ToString() $componentItem.Get_InnerXml().ToString() "{" $currentAnswerNodeValue "}"
                                    $isValid = $false
                                    break
                                }
                                # if match found and not empty and validation Type is "None"
                                # skip this item.
                                elseif ($componentItem.GetAttribute("ValidationType") -eq "None")
                                {
                                    break
                                }

                                #if validation Type is not "None" do validation.
                                $temp = ValidateItem $currentAnswerNodeValue $componentItem.GetAttribute("ValidationType")
                                # if validation failed Log the failure.
                                if ($temp -eq $false)
                                {
                                    Write-Information $true "Validation Failed: " $component.ToString() $componentItem.Get_InnerXml().ToString() "{" $currentAnswerNodeValue $componentItem.GetAttribute("ValidationType") "}"
                                    $isValid = $false
                                }
                                elseif ($temp -eq $true)
                                {
                                    
                                }
                                else
                                {
                                    $isValid = $false
                                    foreach($hashKey in $temp.Get_Keys())
                                    {
                                        if ($temp[$hashKey] -eq $false)
                                        {
                                            Write-Information $true "Validation Failed: " $component.ToString() $componentItem.Get_InnerXml().ToString() "{" $hashKey $componentItem.GetAttribute("ValidationType") "}"
                                        }
                                    }       
                                }
                                write-debug ""
                                break
                            }
                        }
                    }
                }
            }
        }
    }
    return $isValid
}

#############################################################################
# InitializationCheck
# Setup script variables, check arguments etc.
#############################################################################
function InitializationCheck()
{
    write-debug "InitializationCheck"
    # Global variable for storing handle to the clone log file.
    # Change this to Actual Path relative to [C:\Program Files\Microsoft\Exchange Server\]

    $script:success = $false
    $script:uniqueName = "A9ABA4D2-C21C-4bc5-8B30-3EA47BBE3608"
    $script:uniqueKey = "A9ABA4D2C21C4bc58B303EA47BBE3608"
    $script:installPath = GetEdgeInstallPath
    $script:logfilePath = $installPath + "Logging\SetupLogs\cloneLogFile.log"

    $script:logger = new-object System.IO.StreamWriter ($logfilePath , $true)

    if ($key -eq $uniqueKey)
    {
       write-host -ForegroundColor "red" "Warning:Passwords will be encrypted with the default script encryption key"
    }

    # Check/Setup the SecurePassword encryption key 
    if (($key.Length -ne 32) -and
        ($key.Length -ne 24) -and
        ($key.Length -ne 16))
    {
       Write-Information $true "Key Length needs to be 16, 24 or 32 bytes long." 
       return
    }

    $script:encryptionKey = $key.ToCharArray()
    write-debug ("encryption key=" + $encryptionKey)

    Write-LogStartInformation

    # Return error if any extra parameters are supplied.
    if($script:args.Count -gt 0)
    {
        foreach($arg in $script:args)
        {
            Write-Information $true "Invalid additional parameter $arg passed."
        }
        return
    }

    # Adam Service Name
    $AdamServiceName = "ADAM_MSExchange"

    # Make sure ADAM_MsExchange Service is running.
    $adamServiceStatus = get-service $AdamServiceName

    if(($adamServiceStatus -eq $null) -or 
       ($adamServiceStatus.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Running))
    {
        Write-Information $true "Adam service should be running, in order to import clone config."
        return
    }
    
    # Initialize "global" variables that can be used in the template.
    $script:MachineName = [System.Environment]::MachineName

    $script:edgeServer = get-ExchangeServer -Identity:$MachineName | where { $_.IsEdgeServer -eq $true }
    if (!$script:edgeServer)
    {
        Write-Information $true "Please run the Import script on the Edge Server." 
        return
    }

    $isValidInput = validateInput
    if($isValidInput -eq $false)
    {
        return
    }

    $script:success = $true	
}

$cloneTemplate = 
'<cloneTemplate>
<machineSpecificItems>
    <TransportServer Key="None">   
        <Item ValidationType="DirectoryPath">ConnectivityLogPath</Item>            
        <Item ValidationType="DirectoryPath">MessageTrackingLogPath</Item>
        <Item ValidationType="DirectoryPath">PickupDirectoryPath</Item>        
        <Item ValidationType="DirectoryPath">PipelineTracingPath</Item>        
        <Item ValidationType="DirectoryPath">ReceiveProtocolLogPath</Item>
        <Item ValidationType="DirectoryPath">ReplayDirectoryPath</Item>
        <Item ValidationType="DirectoryPath">RoutingTableLogPath</Item>        
        <Item ValidationType="NullableDirectoryPath">RootDropDirectoryPath</Item>                
        <Item ValidationType="DirectoryPath">SendProtocolLogPath</Item>
    </TransportServer>
    <SendConnector Key="Name">
        <Item ValidationType="IPAddress">SourceIPAddress</Item>
    </SendConnector>
    <ReceiveConnector Key="Name">
        <Item ValidationType="Bindings">Bindings</Item>
        <Item ValidationType="FQDN">Fqdn</Item>        
    </ReceiveConnector>
</machineSpecificItems>

<importCloneItems>
    <TransportServer PreCondition="None" ImportVerbs="Set" PostCondition="None">
        <NonClonableItems>
            <Item>Identity</Item>
            <Item>PickupDirectoryDefaultDomain</Item>
            <Item>ExternalDsnReportingAuthority</Item>
            <Item>InternalDsnReportingAuthority</Item>
            <Item>ExternalPostmasterAddress</Item>
            <Item>Name</Item>
        </NonClonableItems> 
    </TransportServer>
    <AcceptedDomain PreCondition="Remove" ImportVerbs="New:Set" PostCondition="Remove">
        <NonClonableItems>
            <Item>Identity</Item>
            <Item>AuthenticationType</Item>
            <Item>LiveIdInstanceType</Item>
        </NonClonableItems> 
    </AcceptedDomain>
    <RemoteDomain PreCondition="Remove" ImportVerbs="New:Set" PostCondition="None">
        <NonClonableItems>
            <Item>Identity</Item>
            <Item>CharacterSet</Item> 
            <Item>NonMimeCharacterSet</Item> 
            <Item>ContentType</Item>
            <Item>LineWrapSize</Item>
            <Item>MsExchRoutingDisplaySenderEnabled</Item>
            <Item>TNEFEnabled</Item>
            <Item>AllowedOOFType</Item>
            <Item>AutoReplyEnabled</Item>
            <Item>AutoForwardEnabled</Item>
            <Item>DeliveryReportEnabled</Item>
            <Item>NDREnabled</Item>
            <Item>MeetingForwardNotificationEnabled</Item>            
            <Item>UseSimpleDisplayName</Item>
            <Item>NDRDiagnosticInfoEnabled</Item>
            <Item>IsInternal</Item>
            <Item>TrustedMailInboundEnabled</Item>
            <Item>TrustedMailOutboundEnabled</Item>
        </NonClonableItems> 
    </RemoteDomain>
    <TransportAgent PreCondition="None" ImportVerbs="Set:EnableOrDisable" PostCondition="None">
    </TransportAgent>
    <ReceiveConnector PreCondition="Remove" ImportVerbs="New" PostCondition="None">
        <NonClonableItems>
            <Item>Server</Item>
            <Item>AdvertisedDomain</Item>
            <Item>PermissionGroups</Item>
        </NonClonableItems> 
    </ReceiveConnector>
    <SendConnector PreCondition="Remove" ImportVerbs="New" PostCondition="None">
    </SendConnector>
    <ContentFilterConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </ContentFilterConfig>
    <SenderIdConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </SenderIdConfig>
    <SenderFilterConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </SenderFilterConfig>
    <RecipientFilterConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </RecipientFilterConfig>
    <AddressRewriteEntry PreCondition="Remove" ImportVerbs="New" PostCondition="None">
    </AddressRewriteEntry>
    <AttachmentFilterEntry PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </AttachmentFilterEntry>
    <AttachmentFilterListConfig PreCondition="SetBlank" ImportVerbs="Set" PostCondition="None">
    </AttachmentFilterListConfig>
    <IPAllowListEntry PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </IPAllowListEntry>
    <IPAllowListProvider PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </IPAllowListProvider>
    <IPAllowListConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </IPAllowListConfig>
    <IPAllowListProvidersConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </IPAllowListProvidersConfig>
    <IPBlockListEntry PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </IPBlockListEntry>
    <IPBlockListProvider PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </IPBlockListProvider>
    <IPBlockListConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </IPBlockListConfig>
    <IPBlockListProvidersConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </IPBlockListProvidersConfig>
    <ContentFilterPhrase PreCondition="Remove" ImportVerbs="Add" PostCondition="None">
    </ContentFilterPhrase>
    <SenderReputationConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
    </SenderReputationConfig>
    <TransportConfig PreCondition="None" ImportVerbs="Set" PostCondition="None">
        <NonClonableItems>
            <Item>Identity</Item>
        </NonClonableItems>
    </TransportConfig>
</importCloneItems>

</cloneTemplate>
'

#######################################################################
# Main Script starts here. Validates the parameters and based on isImport value
# will run through the Import or Create answer file Steps.
#######################################################################

$errorActionSave = $ErrorActionPreference

# Do all the Initialization Checks
InitializationCheck
if (-not $success)
{
    Export-ReleaseHandles
    exit
}

write-debug "Main"

write-debug ("cloneConfigData = $cloneConfigData")

$cloneTemplateDoc.LoadXML($cloneTemplate)

$cloneDataDoc.Load($cloneConfigData)

if($isImport -eq $true)
{
    # Incase of Importing.
    write-verbose "Import Cloned Config"
    if($cloneConfigAnswer -ne $null)
    {
        $cloneAnswerDoc.Load($cloneConfigAnswer)
        trap
        {
            Write-Information $true "Loading Answer File Failed."
            Write-Information $true ("Reason: " + $error[0])
            Export-ReleaseHandles
            exit
        }
    }
    # Incase of Restoring.
    else
    {
        $emptyAnswerString = "<MachineSpecificSettings/>"
        $cloneAnswerDoc = [System.Xml.XmlDocument]($emptyAnswerString)  
    }
    $validationSuccess = ValidateAnswerFile $cloneTemplateDoc $cloneAnswerDoc
    if($validationSuccess -eq $false)
    {
        Write-Information $true ("Validation Process Failed. Please enter correct information in " + $cloneConfigAnswer)
        Export-ReleaseHandles
        exit
    }
    ImportConfig $cloneTemplateDoc $cloneDataDoc $cloneAnswerDoc

    trap
    {
        # Catch password conversion errors
        if ($error[0].Exception.GetType() -eq [System.Security.Cryptography.CryptographicException])
        {
            Write-Information $true "Password Decryption Failed - Check the decryption key."
        }
        else
        {
            Write-Information $true "Importing Edge configuration information Failed."
        }
        Write-Information $true ("Reason: " + $error[0])
        Export-ReleaseHandles
        exit
    }
    Write-Information $true "Importing Edge configuration information Succeeded."
}
else
{
    write-verbose "Validate Cloned Config"

    # Don't display errors For Validation Phase 
    $ErrorActionPreference = "SilentlyContinue"
    
    ValidateAndCreateAnswerFile $cloneTemplateDoc $cloneDataDoc $cloneConfigAnswer
    trap
    {
        # Restore Error Action Preference
        $ErrorActionPreference = $errorActionSave        
    
        Write-Information $true "Creating Answer File Step Failed."
        Write-Information $true ("Reason: " + $error[0])
        Export-ReleaseHandles
        exit
    }
    Write-Information $true ("Answer File is successfully created: " + $cloneConfigAnswer)

}

Export-ReleaseHandles

# SIG # Begin signature block
# MIIabgYJKoZIhvcNAQcCoIIaXzCCGlsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6RZdo4d5DSqeYD1X+EnrL/Jv
# nYigghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggSpMIIEpQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHCMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTt3pnZGNvgXRDzNDVdzyP7mAQMczBiBgorBgEEAYI3AgEMMVQw
# UqAqgCgASQBtAHAAbwByAHQARQBkAGcAZQBDAG8AbgBmAGkAZwAuAHAAcwAxoSSA
# Imh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEB
# BQAEggEAVuHCDUTfHT/GT9IlmwaGqmnKRp+pEOGgZ3frLZrs543pkSpy0SYeU2Fg
# 4AulpkMDmAfrqSZsTkBR5AntGAqfMP4GnAEOhKHi+jy00LQCIdnC29ykP8qhgAKK
# 1dsrclOmbf6DbjicRW2WnAyj4N+6a8R7XL/XA/7VXy4w6uNYlS4+cE+h6aB2pW+C
# s4Ahk/SBOm4MyU/wSbneDrVbUGWIVMNrBh8MAD3R6kJyXXr9ObwWE1ZZL5Ytb2BQ
# b8gy9/JjxXrJJdVOz6jqc9Bxbp0MCZwAunHYlHU6zKrcZVIfdU9pphY+SdqWMSYz
# 3a5RUfVqFq4pBDtQfyZ5uiiXJjpWDKGCAigwggIkBgkqhkiG9w0BCQYxggIVMIIC
# EQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAAArOTJIwbLJSPMA
# AAAAACswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
# KoZIhvcNAQkFMQ8XDTEzMDIwNTA2MzcyM1owIwYJKoZIhvcNAQkEMRYEFP3If+E1
# mFJJEIOpt64Fzje9p9+aMA0GCSqGSIb3DQEBBQUABIIBAF+eFbEa4tUSqFkmwEFO
# RwUFFuuQ2CvLhxfjbMWa3NWmb0n0GdS2YShJVlTG9dJhHT8OC47zaRYkvt/GdnnX
# 4DRgfkUX9TOUejl66jRMjnpGQrwct2Eye5oloB8RBZqdV/gFuAYTC6MrzLeGvPEr
# jc028iVTPk15/qIvO5qapG+hX1hHxNopMmDakt7D2rH4WEbdFPZTf8xthuX+xr5M
# 67BkryhunbfawBhEHW444ZwdmE/0KHq8ljXyIeb3vTXW1Poj/OM9/asnT3xBMhvy
# ABhXlrKRv8fxg0ELRikNuXZWW6/33druTV5ToRz3q3kOCUxsnD6QVNkOe1CgUzbY
# +Qs=
# SIG # End signature block
