##################################################################### 
#Script Title: Get Active Directory Object Category Information Tool   
#Script File Name: Get-ADObjectCategoryInfo.ps1  
#Author: Ron Ratzlaff (aka "The_Ratzenator")   
#Date Created: 9/30/2014   
##################################################################### 
   
#Requires -Version 3.0 
#Requires -Module ActiveDirectory 
 
 
 
 
Function Get-ADObjectCategoryInfo 
{ 
    <#  
      .SYNOPSIS  
        
          Retrieves specified Active Directory object category information 
          
      .DESCRIPTION  
        
          The "Get Active Directory Object Category Information Tool” utilizes the Get-ADObject cmdlet and uses its filter parameter to target the ObjectCategory attribute. This tool retrieves information from the following five objects within the ObjectCategory attribute: Computer, Domain, Group, Organizational Unit, and User. This tool can retrieve the following five properties of each Active Directory Object Category object: Canonical Names, Distinguished Names, Names, Object Classes, and Object GUIDS. This tool can also produce a count of the total number of specified Active Directory Object Category objects. The ShowAll parameter will retrieve the count and all five properties of a specified Active Directory Object Category object and it cannot be used simultaneously with other parameter. Since the ShowAll parameter retrieves all of the information that the other parameters would retrieve, it would be redundant and superfluous to use the other parameters as well.  
        
      .EXAMPLE  
  
        Retrieve all five properties including a count for one of the five Active Directory Object Category objects in a compacted formatted list:  
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowAll 
 
        Or  
 
        Get-ADObjectCategoryInfo -Object "Computer" -SA 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowAll 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SA 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowAll 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SA 
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowAll 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SA 
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowAll 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "User" -SA 
  
      .EXAMPLE  
          
        Retrieve all five properties including a count for one of the five Active Directory Object Category objects with headers instead of a compacted formatted list: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowCount -ShowCanonicalName -ShowDistinguishedName -ShowName -ShowObjectClass -ShowObjectGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SC -SCN -SDN -SN -SOC -SOG 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowCount -ShowCanonicalName -ShowDistinguishedName -ShowName -ShowObjectClass -ShowObjectGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SC -SCN -SDN -SN -SOC -SOG 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowCount -ShowCanonicalName -ShowDistinguishedName -ShowName -ShowObjectClass -ShowObjectGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SC -SCN -SDN -SN -SOC -SOG 
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowCount -ShowCanonicalName -ShowDistinguishedName -ShowName -ShowObjectClass -ShowObjectGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SC -SCN -SDN -SN -SOC -SOG 
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowCount -ShowCanonicalName -ShowDistinguishedName -ShowName -ShowObjectClass -ShowObjectGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SC -SCN -SDN -SN -SOC -SOG 
  
      .EXAMPLE  
  
        Retrieve just the count for one of the five Active Directory Object Category object: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowCount  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SC 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowCount 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SC 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowCount 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SC 
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowCount  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SC  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowCount  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SC  
      
      .EXAMPLE  
          
        Retrieve just the Canonical Name for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowCanonicalName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SCN 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowCanonicalName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SCN 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowCanonicalName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SCN  
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowCanonicalName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SCN  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowCanonicalName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SCN 
  
      .EXAMPLE  
  
        Retrieve just the Distinguished Name for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SDN 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SDN 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SDN  
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SDN  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SDN   
  
      .EXAMPLE  
          
        Retrieve just the Name for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowName 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SN 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SN  
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SN  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SN  
  
      .EXAMPLE  
  
        Retrieve just the Object Class for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowObjectClass  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SOC 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowObjectClass  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SOC 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowObjectClass  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SOC 
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowObjectClass  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SOC  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowObjectClass  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SOC  
      
      .EXAMPLE  
          
        Retrieve just the Name for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowName 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SN 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SN  
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SN  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SN  
  
      .EXAMPLE  
  
        Retrieve just the Object GUID for one of the five Active Directory Object Category objects: 
  
        AD Object: Computer 
        ------------------- 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowObjectCGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -SOC 
         
        AD Object: Domain 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Domain" -ShowObjectCGUID 
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Domain" -SOC 
 
        AD Object: Group 
        ----------------- 
         
        Get-ADObjectCategoryInfo -Object "Group" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "Group" -SOC 
 
        AD Object: Organizational Unit 
        ------------------------------ 
         
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "OrganizationalUnit" -SOC  
 
        AD Object: User 
        --------------- 
         
        Get-ADObjectCategoryInfo -Object "User" -ShowDistinguishedName  
 
        or 
 
        Get-ADObjectCategoryInfo -Object "User" -SOC  
        
      .EXAMPLE  
          
        Mix and match property parameters for specified Active Directory Object Category objects: 
  
        Get-ADObjectCategoryInfo -Object "Computer" -ShowCount -ShowCanonicalName  
         
        Or 
         
        Get-ADObjectCategoryInfo -Object "Computer" -ShowCanonicalName -ShowDistinguishedName 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowDistinguishedName -ShowName 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowName -ShowObjectClass 
 
        Or 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowObjectClass -ShowObjectGUID 
 
      .EXAMPLE 
 
        Output results simultaneously to screen and to a file: 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowAll | Tee-Object -FilePath "$env:TEMP\Get-ADObjectCategoryInfo.log" 
 
      .EXAMPLE 
 
        Output results to a file only: 
 
        Get-ADObjectCategoryInfo -Object "Computer" -ShowAll | Out-File -FilePath "$env:TEMP\Get-ADObjectCategoryInfo.log" 
  
      .PARAMETER Object  
        
          A mandatory parameter used to specifiy the Active Directory Object Category object.  
        
      .PARAMETER ShowAll  
        
          A parameter that retrieves all five properties including a count of the total number of specified Active Directory Objcet Category objects and displays the results in a compact formatted list.  
        
      .PARAMETER ShowCount 
        
          A parameter used to count the total number of specfied Active Directory Object Category objects in the Active Directory environment.   
 
      .PARAMETER ShowCanonicalName 
        
          A parameter used to retrieve all Canonical Names within the Active Directory enviornment.   
 
      .PARAMETER ShowDistinguishedName 
        
          A parameter used to retrieve all Distinguished Names within the Active Directory enviornment. 
 
      .PARAMETER ShowName 
        
          A parameter used to retrieve all Names within the Active Directory enviornment. 
 
      .PARAMETER ShowObjectClass 
        
          A parameter used to retrieve all Object Classes within the Active Directory enviornment. 
 
      .PARAMETER ShowObjectGUID 
        
          A parameter used to retrieve all Object GUIDs within the Active Directory enviornment 
      #>  
       
      [cmdletbinding()]  
      
        Param   
        (               
             [Parameter(Mandatory=$true, 
               Position = 0, 
               HelpMessage='What AD object would you like to target?')]   
             [ValidateSet('Computer', 'Domain', 'Group', 'OrganizationalUnit', 'User', IgnoreCase = $true)]    
             $Object, 
          
             [Parameter(   
                HelpMessage='Do you want to display all of the properties for AD objects?')]    
             [Alias('SA')] 
             [switch]$ShowAll, 
          
             [Parameter( 
                HelpMessage='Do you want to show the canonical name of AD objects?')]     
             [Alias('SCN')] 
             [switch]$ShowCanonicalName, 
          
             [Parameter(  
                HelpMessage='Do you want to show the count of AD objects?')]    
             [Alias('SC')] 
             [switch]$ShowCount, 
 
             [Parameter( 
                HelpMessage='Do you want to show the distinguished name of AD objects?')]    
             [Alias('SDN')] 
             [switch]$ShowDistinguishedName, 
          
             [Parameter( 
                HelpMessage='Do you want to show the name of AD objects?')]    
             [Alias('SN')] 
             [switch]$ShowName, 
 
             [Parameter( 
                HelpMessage='Do you want to show the object class of AD objects?')]    
             [Alias('SOC')] 
             [switch]$ShowObjectClass, 
 
             [Parameter( 
                HelpMessage='Do you want to show the object GUID of AD objects?')]    
             [Alias('SOG')] 
             [switch]$ShowObjectGUID 
        )    
 
    Begin 
    { 
        #region ----------Declare Get-ADObject variables---------- 
         
        $NewLine = "`r`n" 
         
        $ADObjectCanonicalName = (Get-ADObject -Filter 'ObjectCategory -eq $Object' -Properties CanonicalName).CanonicalName 
         
        $ADObjectClass = (Get-ADObject -Filter 'ObjectCategory -eq $Object' -Properties ObjectClass).ObjectClass 
         
        $ADObjectDistinguishedName = (Get-ADObject -Filter 'ObjectCategory -eq $Object' -Properties DistinguishedName).DistinguishedName 
 
        $ADObjectGUID = (Get-ADObject -Filter 'ObjectCategory -eq $Object' -Properties ObjectGUID).ObjectGUID  
 
        $ADObjectName = (Get-ADObject -Filter 'ObjectCategory -eq $Object').Name  
         
        $ADObjectShowAll = Get-ADObject -Filter 'ObjectCategory -eq $Object' -Properties CanonicalName, DistinguishedName, ObjectClass, ObjectGUID  
         
        #endregion ----------Declare Get-ADObject variables---------- 
         
         
        #region ----------Create a variable to contain a hashtable for counting AD objects---------- 
         
        $GetCount = @( 
 
                If ($Object -eq "Computer") 
                { 
                    $NewLine 
                         
                    Write-Output "Count" 
                    Write-Output "-----" 
 
                    $NewLine 
                             
                    $GetADObject = Get-ADObject -Filter 'ObjectCategory -eq "Computer"' 
 
                    $GetADObject.Count 
 
                    $NewLine 
                } 
 
                ElseIf ($Object -eq "Domain") 
                { 
                    $NewLine 
                         
                    Write-Output "Count" 
                    Write-Output "-----" 
 
                    $NewLine 
                             
                    $GetADObject = Get-ADObject -Filter 'ObjectCategory -eq "Domain"' 
 
                    $GetADObject.Name.Count 
 
                    $NewLine 
                } 
 
                ElseIf ($Object -eq "Group") 
                { 
                    $NewLine 
                         
                    Write-Output "Count" 
                    Write-Output "-----" 
 
                    $NewLine 
                             
                    $GetADObject = Get-ADObject -Filter 'ObjectCategory -eq "Group"' 
 
                    $GetADObject.Count 
 
                    $NewLine 
                } 
 
                ElseIf ($Object -eq "OrganizationalUnit") 
                { 
                    $NewLine 
                         
                    Write-Output "Count" 
                    Write-Output "-----" 
 
                    $NewLine 
                             
                    $GetADObject = Get-ADObject -Filter 'ObjectCategory -eq "OrganizationalUnit"' 
 
                    $GetADObject.Count 
 
                    $NewLine 
                } 
 
                ElseIf ($Object -eq "User") 
                { 
                    $NewLine 
                         
                    Write-Output "Count" 
                    Write-Output "-----" 
 
                    $NewLine 
                             
                    $GetADObject = Get-ADObject -Filter 'ObjectCategory -eq "User"' 
 
                    $GetADObject.Count 
 
                    $NewLine 
                }) 
 
        #endregion ----------Create a variable to contain a hashtable for counting AD objects---------- 
                 
         
        #region ----------Create variables to contain hashtables for determining parameter count and multiple parameter usage---------- 
         
        $ParamCount = @( 
         
                $Params = @(      
                                Foreach ($Param in 'ShowAll', 'ShowCanonicalName', 'ShowCount', 'ShowDistinguishedName', 'ShowName', 'ShowObjectClass', 'ShowObjectGUID') 
                                { 
                                    If ($PSBoundParameters[$Param])  
                                    {  
                                        $Param  
                                    } 
                                } 
                            ) 
 
                    Write-Output  $Params.Count) 
 
        $ParamString = @( 
                 
                $Params = @(      
                                Foreach ($Param in 'ShowAll', 'ShowCanonicalName', 'ShowCount', 'ShowDistinguishedName', 'ShowName', 'ShowObjectClass', 'ShowObjectGUID') 
                                { 
                                    If ($PSBoundParameters[$Param])  
                                    {  
                                        $Param 
                                    } 
                                } 
                            ) 
 
                    Write-Output $Params) 
 
        #endregion ----------Create variables to contain hashtables for determining parameter count and multiple parameter usage---------- 
         
         
        #region ----------Create variables to contain hashtables for determining multiple parameter usage---------- 
         
        If ($ParamCount -eq "7") 
        { 
            $NewLine 
                     
            Write-Warning "The ShowAll parameter cannot be used simultaneously with other parameters" 
 
            $NewLine 
 
            Write-Output "If you want to display all of the $Object results, then use the ShowAll parameter exclusively (Ex: Get-ADObjectCategoryInfo -Object $Object -ShowAll)" 
 
            $NewLine 
     
            Write-Output "The script will now exit. Please rerun the script with the appropriate parameter(s)" 
                 
            $NewLine 
 
            Pause 
 
            $NewLine 
 
            Exit  
        } 
 
        ElseIf ($ParamCount -gt "1" -and $ParamString -match "ShowAll") 
        { 
            $NewLine 
                     
            Write-Warning "The ShowAll parameter cannot be used simultaneously with other parameters" 
 
            $NewLine 
 
            Write-Output "If you want to display all of the $Object results, then use the ShowAll parameter exclusively (Ex: Get-ADObjectCategoryInfo -Object $Object -ShowAll)" 
 
            $NewLine 
     
            Write-Output "The script will now exit. Please rerun the script with the appropriate parameter(s)" 
                 
            $NewLine 
 
            Pause 
 
            $NewLine 
 
            Exit  
        } 
 
        $SwitchParamString = @( 
 
                Switch ($ParamString) 
                { 
                    "ShowAll" 
                    { 
                        $NewLine 
                         
                        Write-Output "Show All" 
                        Write-Output "--------" 
 
                        $GetCount 
                         
                        $ADObjectShowAll 
 
                        $NewLine  
                    } 
                     
                    # Begin processing ShowCount combinations 
 
                    {   
                        $_ -eq "ShowCount" -and $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" -and $_ -eq "ShowGUID"  
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowCanonicalName" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowDistinguishedName" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowName" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine  
                    } 
 
                    { 
                        $_ -eq "ShowCount" -and $_ -eq "ShowObjectGUID" 
                    } 
 
                    { 
                        $GetCount 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine  
                    } 
 
                    "ShowCount" 
                    { 
                        $GetCount  
                    } 
 
                    # Begin processing ShowCanonicalName combinations 
 
                    {   
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" -and $_ -eq "ShowGUID"  
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowDistinguishedName" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowName" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowCanonicalName" -and $_ -eq "ShowObjectGUID" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    "ShowCanonicalName" 
                    { 
                        $NewLine 
 
                        Write-Output "Canonical Names" 
                        Write-Output "---------------" 
 
                        $NewLine 
 
                        $ADObjectCanonicalName 
                    } 
 
                    # Begin processing ShowDistinguishedName combinations 
 
                    {   
                        $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" -and $_ -eq "ShowGUID"  
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowName" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowDistinguishedName" -and $_ -eq "ShowObjectGUID" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    "ShowDistinguishedName" 
                    { 
                        $NewLine 
 
                        Write-Output "Distinguished Names" 
                        Write-Output "-------------------" 
 
                        $NewLine 
 
                        $ADObjectDistinguishedName 
                    } 
 
                    # Begin processing ShowName combinations 
 
                    {   
                        $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" -and $_ -eq "ShowGUID"  
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowName" -and $_ -eq "ShowObjectClass" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    { 
                        $_ -eq "ShowName" -and $_ -eq "ShowObjectGUID" 
                    } 
 
                    { 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    "ShowName" 
                    { 
                        $NewLine 
 
                        Write-Output "Names" 
                        Write-Output "-----" 
 
                        $NewLine 
 
                        $ADObjectName 
 
                        $NewLine 
                    } 
 
                    # Begin processing ShowObjectClass combinations 
 
                    {   
                        $_ -eq "ShowObjectClass" -and $_ -eq "ShowGUID"  
                    } 
 
                    { 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
 
                    "ShowObjectClass" 
                    { 
                        $NewLine 
 
                        Write-Output "Object Classes" 
                        Write-Output "--------------" 
 
                        $NewLine 
 
                        $ADObjectClass 
 
                        $NewLine 
                    } 
 
                    # Begin processing ShowObjectGUID combinations 
                     
                    "ShowObjectGUID" 
                    { 
                        $NewLine 
 
                        Write-Output "Object GUIDs" 
                        Write-Output "------------" 
 
                        $NewLine 
 
                        $ADObjectGUID 
 
                        $NewLine 
                    } 
                }) 
 
        #endregion ----------Create variables to contain hashtables for determining multiple parameter usage---------- 
    } 
 
    Process 
    { 
        Switch ($Object) 
        { 
            "Computer" 
            { 
                $NewLine 
 
                Write-Output "AD Object: Computer" 
                Write-Output "-------------------" 
                  
                $SwitchParamString 
            } 
 
            "Domain" 
            { 
                $NewLine 
 
                Write-Output "AD Object: Domain" 
                Write-Output "-----------------" 
                  
                $SwitchParamString 
            } 
 
            "Group" 
            { 
                $NewLine 
 
                Write-Output "AD Object: Group" 
                Write-Output "----------------" 
                  
                $SwitchParamString 
            } 
 
            "OrganizationalUnit" 
            { 
                $NewLine 
 
                Write-Output "AD Object: Organizational Unit" 
                Write-Output "------------------------------" 
                  
                $SwitchParamString 
            } 
 
            "User" 
            { 
                $NewLine 
 
                Write-Output "AD Object: User" 
                Write-Output "---------------" 
                  
                $SwitchParamString 
            } 
        } 
    }            
 
    End {} 
}