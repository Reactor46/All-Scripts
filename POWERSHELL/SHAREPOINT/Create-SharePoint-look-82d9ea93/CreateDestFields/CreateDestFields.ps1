# check to ensure Microsoft.SharePoint.PowerShell is loaded
Write-Host "Loading SharePoint Powershell Snapin" -ForegroundColor Blue
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue

#region Glable Variable
[int]$global:count = 0;
#endregion Glable Variable

#region Add-List-Content-Main-Function

function Add-List-Content-Main-Function ([string] $configFile)
{
	$AdminServiceName = "SPAdminV4";
	$IsAdminServiceRunning = $true;

	Add-PSSnapin Microsoft.SharePoint.PowerShell –ErrorAction SilentlyContinue	;
	
	cls;
	if ($(Get-Service $AdminServiceName).Status -eq "Stopped")
	{
	    $IsAdminServiceRunning = $false;
	    Start-Service $AdminServiceName       
	}
	
	if([string]::IsNullOrEmpty($configFile))
	{
		return
	}
	[xml]$solutionsConfig = Get-Content $configFile

	if ($solutionsConfig -eq $null)
	{
		return
	}
	#Log File Format Function
	#$LogFileDayF = Get-Log-File-Name;
	
	[string] $SourceSiteUrl = ""
	[string] $DestinationSiteUrl = ""
	
	#[int] $count = 0
	
	$solutionsConfig.ParentNode | forEach-Object{        
		$SourceSiteUrl = $_.SourceSiteUrl
		$DestinationSiteUrl = $_.DestinationSiteUrl
	}

    foreach ($List in $solutionsConfig.ParentNode.ListNames)
	{
	 write-Host $List.ListName.SourceListName;
	}

	$solutionsConfig.ParentNode.ListNames | forEach-Object{		
		[string]$SourceListName  = $_.ListName.SourceListName;
		[string]$DestinationListName  = $_.ListName.DestinationListName;
		
		[string]$PrimaryKeyValue  = $_.ListName.PrimaryKeyValue;		
		[string]$FieldNameForFltrCondValue  = $_.FilterCondForADILog.FilterFieldName;
		[string]$FieldTypeForFltrCondValue  = $_.FilterCondForADILog.FilterFieldType;
		[string]$KeyValueForFltrCondValue  = $_.FilterCondForADILog.FilterKeyValue;
		[string]$DestinationLookUpList  = $_.LookUPDetails.LookupListNameDest;
		
		#[int] $countField = 0
		#$solutionsConfig.ParentNode.ListNames.AddFields | forEach-Object{
		
        foreach ($Field in $_.AddFields.AddField)
		{
	 	write-Host $Field.DisplayName;
		}
		
		#$solutionsConfig.ParentNode.ListNames.AddFields.AddField | forEach-Object{
		foreach ($Field in $_.AddFields.AddField)
		{
		
		[string]$fieldDisplayName  = $Field.DisplayName;
		[string]$fieldInternalName  = $Field.InternalName;
		[string]$fieldDataType  = $Field.DataType;
		
		[string]$LookUpMasterListName  = $Field.LookUpMasterListName;
		[string]$LookUpMasterListClmnName  = $Field.LookUpMasterListClmnName;
		
		[string]$MultiLineClmnTextType  = $Field.MultiLineClmnTextType;
		
		[string]$ChoiceClmnDisplayType  = $Field.ChoiceClmnDisplayType;
		[string]$ChoiceClmnItems  = $Field.ChoiceClmnItems;
		
		[string]$CalculatedClmnFormula  = $Field.CalculatedClmnFormula;
		[string]$CalculatedClmnDataType  = $Field.CalculatedClmnDataType;
		
		try
		{
		write-Host $fieldDisplayName;
		#CreateField $fieldDisplayName $fieldInternalName $fieldDataType $DestinationSiteUrl $DestinationListName $LookUpMasterListName $LookUpMasterListClmnName $MultiLineClmnTextType $ChoiceClmnDisplayType $ChoiceClmnItems $CalculatedClmnFormula $CalculatedClmnDataType;
		
		}
		catch
		{
			$ErrMsg = "Error : While adding files/document(s) and error is $_"
			Write-Host -ForegroundColor red $ErrMsg;
		}
		}
		
		#}
		
		
		[string]$SourceFieldName  = $_.FieldNames.source.FieldName
		[string]$DestinationFieldName  = $_.FieldNames.destination.FieldName
										
		if ($SourceFieldName -eq "")
		{			 					   
			Write-Host "Source Field Name is null";
			return;
		}
		elseif ($DestinationFieldName -eq "")
		{			 					   
			Write-Host "DestinationFieldName is null";
			return;
		}
		else
		{
#				#region function is for generating/appending Log File
#				Generate-LogFile "" $count $LogFileDayF;				
#				#endregion function is for generating/appending Log File
#								
#				$keyValueSourceFields = $SourceFieldName.split(",")	
#				$keyValueDesinationFields = $DestinationFieldName.split(",")	
#				
#				if($keyValueSourceFields.length -ne $keyValueDesinationFields.length)
#				{
#					$ErrMsg = "Error : Field(s) count of source list $SourceListName and destination list $DestinationListName does not match.Kindly check.";
#					Write-Host -BackgroundColor DarkBlue -ForegroundColor Yellow $ErrMsg;
#					
#					#region function is for generating/appending Log File
#					#Bellow function is for generating Log File		
#					$count=$count+1;
#					Generate-LogFile $ErrMsg $count $LogFileDayF;
#					#endregion function is for generating/appending Log File
#				}
#				else
#				{				
#					#Add-List-Content $SourceSiteUrl $DestinationSiteUrl $SourceListName $DestinationListName $SourceFieldName $DestinationFieldName $PrimaryKeyValue $FieldNameForFltrCondValue $FieldTypeForFltrCondValue $KeyValueForFltrCondValue $LogFileDayF $DestinationLookUpList;					
#				}
				
		}		
	}
	
	if (-not $IsAdminServiceRunning)
	{
    	Stop-Service $AdminServiceName     
	}
	
		Write-Host "Removing SharePoint Powershell Snapin" -ForegroundColor Blue;
		Remove-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue;
		$count = 0;
		Write-Host "Finished...."  -ForegroundColor Blue
}

#endregion Add-List-Content-Main-Function

function CreateField([string] $fieldDisplayName,[string] $fieldInternalName,[string] $fieldDataType, [string] $DestinationSiteUrl,
                     [string] $DestinationListName, [string] $LookUpMasterListName, [string] $LookUpMasterListClmnName, [string] $MultiLineClmnTextType,
					 [string] $ChoiceClmnDisplayType, [string] $ChoiceClmnItems, [string] $CalculatedClmnFormula, [string] $CalculatedClmnDataType )
{
    # Load Web and List 
	try{
	
	
	
	#region Load List and Web
	$dstListWeb = Get-SPWeb -identity $DestinationSiteUrl;	
	$destinationListUrl = $dstListWeb.ServerRelativeUrl + "/" + $DestinationListName;
	$DestinationList = $dstListWeb.GetList($destinationListUrl);
	#endregion
	
    #region Create Field
	
    if($fieldDataType -eq "Text")
	{
	#region column of type 'Single line of text'
    $ColXml = "<Field Type='Text' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
    #endregion
	#$DestinationList.Fields.Add("CalcField", "Calculated", 0)
	}
	elseif($fieldDataType -eq "Note")
	{
	#region column of type 'Multi line of text'
	if($MultiLineClmnTextType -eq "PlainText")
	{
    $ColXml = "<Field Type='Note' NumLines='6' RichText='FALSE' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";
    }
	elseif($MultiLineClmnTextType -eq "RichText")
	{
    $ColXml = "<Field Type='Note' NumLines='6' RichText='TRUE' RichTextMode='Compatible' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";
    }
	elseif($MultiLineClmnTextType -eq "EnhancedRichText")
	{
    $ColXml = "<Field Type='Note' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' IsolateStyles='TRUE' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";
    }
	
	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "Choice")
	{
	#region column of type 'Choice'
	
	#region Get Choice Items
	$ChoiceItems = $ChoiceClmnItems.split(";");
	$Choices = "<CHOICES>";
	#[System.Text.StringBuilder] $Choices = New-Object [System.Text.StringBuilder];
	#$Choices.Append("<CHOICES>");
	for ($k=0; $k -le $ChoiceItems.length; $k++)
		{
		$Choices = $Choices + "<CHOICE>";
		$Choices = $Choices + $ChoiceItems[$k];
		$Choices = $Choices + "</CHOICE>";
		}
	$Choices = $Choices + "</CHOICES>";
	
	#endregion
	
	if($ChoiceClmnDisplayType -eq "Dropdown")
	{
    $ColXml = "<Field Type='Choice' Format='Dropdown' FillInChoice='FALSE' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' >" + $Choices + "</Field>";
    }
	elseif($ChoiceClmnDisplayType -eq "RadioButtons")
	{
    $ColXml = "<Field Type='Choice' Format='RadioButtons' FillInChoice='FALSE' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' >" + $Choices + "</Field>";
    }
	elseif($ChoiceClmnDisplayType -eq "MultiChoice")
	{
    $ColXml = "<Field Type='Choice' Format='MultiChoice' FillInChoice='FALSE' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' >" + $Choices + "</Field>";
    }	
    
	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "Number")
	{
	#region column of type 'Number'
	$ColXml = "<Field Type='Number' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "DateTime")
	{
	#region column of type 'DateTime'
	$ColXml = "<Field Type='DateTime' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' />";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "Calculated")
	{
	#region column of type 'Calculated'
	#$ColXml = "<Field Type='Calculated' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "'Formula='"+ $CalculatedClmnFormula +"' />";

	#$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	#$DestinationList.Update();
	$DestinationList.Fields.Add($fieldDisplayName,"Calculated",0);
	$SPField = $DestinationList.Fields.GetField($fieldDisplayName);
	$SPField.Formula = $CalculatedClmnFormula;
	$SPField.Update();
	#endregion
	}
	elseif($fieldDataType -eq "Lookup")
	{
	#region column of type 'Lookup'	
	$lookupListUrl = $dstListWeb.ServerRelativeUrl + "/Lists/" + $LookUpMasterListName;
	$LookupList = $dstListWeb.GetList($lookupListUrl);
	$LookupListID = $LookupList.ID;
	
	$ColXml = "<Field Type='Lookup' RelationshipDeleteBehavior='None' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' List='{" + $LookupListID + "}' ShowField='" + $LookUpMasterListClmnName +"' />";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "Boolean")
	{
    #region column of type 'Boolean'
	$ColXml = "<Field Type='Boolean' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' ><Default>1</Default></Field>";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "User")
	{
	#region column of type 'User'
	$ColXml = "<Field Type='User' List='UserInfo' ShowField='ImnName' UserSelectionMode='PeopleOnly' UserSelectionScope='0' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' ></Field>";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	elseif($fieldDataType -eq "URL")
	{
	#region column of type 'URL'
	$ColXml = "<Field Type='URL' Format='Hyperlink' DisplayName='" + $fieldDisplayName + "' StaticName='" + $fieldInternalName +"' Name='" + $fieldInternalName + "' ></Field>";

	$DestinationList.Fields.AddFieldAsXml($ColXml,$true, [Microsoft.SharePoint.SPAddFieldOptions]::AddFieldToDefaultView);
	$DestinationList.Update();
	#endregion
	}
	#endregion
	
    }
	catch{
			$ErrMsg = "Error : $_"
			Write-Host -ForegroundColor red $ErrMsg;
	}
}

Add-List-Content-Main-Function "C:\Users\Administrator\Desktop\Moving_ListAndDocumentLib_Items\Moving_ListAndDocumentLib_Items\MovingContentConfigFile.xml"