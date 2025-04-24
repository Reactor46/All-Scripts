Workflow Update-SharePointListItem

{
    Param(
    [Parameter(Mandatory=$true)][String]$SharepointSiteURL,
    [Parameter(Mandatory=$true)][String]$SavedCredentialName,
    [Parameter(Mandatory=$true)][String]$ListName,
    [Parameter(Mandatory=$true)][String]$ListItemID,
    [Parameter(Mandatory=$true)][String]$PropertyName,
    [Parameter(Mandatory=$true)][String]$PropertyValue
    )
    
    Function Update-SP2013ListItem
    {
        Param([String]$SiteUri, [String]$itemURI, [String]$PropertyName, [String]$PropertyValue, [string]$ListItemEntityTypeFullName, [PSCredential]$credential)

        $ContextInfoUri = "$SiteUri`/_api/contextinfo"
        $RequestDigest = (Invoke-RestMethod -Method Post -Uri $ContextInfoUri -Credential $credential).GetContextWebInformation.FormDigestValue
        $body = "{ '__metadata': { 'type': '$ListItemEntityTypeFullName' }, $PropertyName`: '$PropertyValue'}"
        $header = @{
            "accept" = "application/json;odata=verbose"
            "X-RequestDigest" = $RequestDigest
            "If-Match"="*"
        }
        Try {
            Invoke-RestMethod -Method MERGE -Uri $itemURI -Body $body -ContentType "application/json;odata=verbose" -Headers $header -Credential $credential
            $Updated = $true
        } Catch {
            $Updated = $false
        }
        $Updated
    }

    # Get the credential to authenticate to the SharePoint List with

    $credential = Get-AutomationPSCredential -Name $SavedCredentialName

    # combined uri

    #SharePoint 2013
    $ListItemsUri = [System.String]::Format("{0}/_api/web/lists/getbytitle('{1}')/items",$SharepointSiteURL, $ListName)
    $ListUri = [System.String]::Format("{0}/_api/web/lists/getbytitle('{1}')",$SharepointSiteURL, $ListName)

    #Get ListItemEntityTypeFullName
    $List = Invoke-RestMethod -uri $ListUri -credential $Credential
    $ListItemEntityTypeFullName = $list.entry.content.properties.ListItemEntityTypeFullName
    $ListItemEntityTypeFullName

    #Translating Field display name (title) to the internal name
    $FieldFilter = "Title eq '$PropertyName'"
    $ListFieldUri = "$ListUri`/Fields?`$Filter=$FieldFilter"
    $ListField = Invoke-RestMethod -Uri $ListFieldUri -Credential $credential
    $FieldInternalName = $ListField.Content.properties.InternalName

    #Get list items
    $listItemURI = inlinescript {
        $listItems = Invoke-RestMethod -Uri $Using:ListItemsUri -Credential $Using:credential
        foreach($li in $listItems)
        {
            $ItemId = $li.Id.split("/")[2].replace("Items", "")
            $ItemId = $ItemId.replace("(","")
            $ItemId = $ItemId.replace(")","")
            If ($ItemId -eq $USING:ListItemID)
            {
                #This is the item URI for the specific list item that we are looking for.
                $itemUri = [System.String]::Format("{0}/_api/{1}",$USING:SharepointSiteURL, $li.id)
            }
        }
        $itemUri
    }
    #Update the list property
    If ($listItemURI)
    {
        Write-Output "Updating $listItemURI. Setting $PropertyName to '$PropertyValue'"
        $Updated = Update-SP2013ListItem -SiteUri $SharepointSiteURL -itemURI $listItemURI -PropertyName $FieldInternalName -PropertyValue $PropertyValue -ListItemEntityTypeFullName $ListItemEntityTypeFullName -credential $credential
    }

    If ($Updated -eq $true)
    {
        Write-OutPut "List item $listItemURI successfully updated."
    } else {
        Write-Error "Failed to update the list item $listItemURI."
    }
} 