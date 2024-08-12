###############################################################################################
##           Author: Vikas Sukhija                   
##           Date: 07-10-2012 
##           Updated: 03/04/2018
##           modified : 06-08-2013 (made it folder independent)
##           Updated : To be used with AD module.
##           Description:- This script is used for feeding data to AD attributes 
##userid,Email,Physicaladdress,City,State,Zip,Country,CountryCode,OfficeTelephone                                                             
###############################################################################################

# Add AD module

Import-Module ActiveDirectory

# Import CSV file that is populated with user id & email address

$now = Get-Date -format "dd-MMM-yyyy HH:mm"

# replace : by -

$now = $now.ToString().Replace(":", "-")

$data = import-csv $args[0]

# Loop thru the data from CSV

foreach ($i in $data)
{
	
	$userid = $i.userid
	$iaddress = $i.PhysicalAddress.Trim()
	$iCity = $i.City.Trim()
	$iCountry = $i.Country.Trim()
	$iCountryCode = $i.CountryCode.Trim()
	$iZip = $i.Zip.Trim()
	$istate = $i.State.Trim()
	$ioffice = $i.OfficeTelephone.Trim()
	
	$userobject = get-aduser -id $userid -Properties "Co", "C", "StreetAddress", "City", "PostalCode", "State", "telephoneNumber"
	
	$Address = $userobject.StreetAddress
	$City = $userobject.City
	$Country = $userobject.co
	$Zip = $userobject.PostalCode
	$State = $userobject.State
	$Phone = $userobject.telephoneNumber
	
	$Log = ".\logs\" + "Currentstate" + $now + ".log"
	
	Add-Content $Log "------------------"
	Add-Content $Log "UserID:$userid"
	Add-Content $Log "Address:$Address"
	Add-Content $Log "City:$City"
	Add-Content $Log "Country:$Country"
	Add-Content $Log "Zip:$Zip"
	Add-Content $Log "State:$State"
	Add-Content $Log "Phone:$Phone"
	Add-Content $Log "------------------"
	############################StreetAddress########################################################
	
	# adding log to check current  address is blank
	
	if ($Address -like $null)
	{
		
		Write-host "$userid has blank Physical address"
		$Log1 = ".\logs\" + "BlankAddress" + $now + ".log"
		Add-content  $Log1 "$userobject has blank Physical address"
		# If address is Blank then populate the address from the csv file
		try
		{
			$Log3 = ".\logs\" + "SetAddress" + $now + ".log"
			Set-ADUser -id $userid -StreetAddress $iaddress
			Add-content  $Log3 "For $userid $iaddress as address has been set"
			Write-Host "For $userid $iaddress as address has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $iaddress Exception occured"
		}
	}
	else
	{
		# adding log to check current address is not blank , then address will be overwritten.		
		try
		{
			$Log2 = ".\logs\" + "Currentaddress" + $now + ".log"
			Set-ADUser -id $userid -StreetAddress $iaddress
			Add-content  $Log2 "For $userid $iaddress as address has been OverWritten"
			Write-host "$userid has been overwritten with $iaddress as Current Physical address"
		}
		catch
		{
			Add-content  $Log2 "For $userid $iaddress Exception occured"
		}
		
	}
	###############################################City#####################################################
	if ($City -like $null)
	{
		Write-host "$userid has blank City"
		$Log1 = ".\logs\" + "BlankCity" + $now + ".log"
		Add-content  $Log1 "$userid has blank City"
		try
		{
			$Log3 = ".\logs\" + "SetCity" + $now + ".log"
			Set-ADUser -id $userid -City $iCity
			Add-content  $Log3 "For $userid $iCity as city has been set"
			Write-Host "For $userid $iCity as city has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $iCity Exception occured"
		}
	}
	else
	{
		
		try
		{
			$Log2 = ".\logs\" + "CurrentCity" + $now + ".log"
			Set-ADUser -id $userid -City $iCity
			Add-content  $Log2 "For $userid $iCity as Current City"
			Write-host "$userid has been overwritten with $iCity as Current City has been OverWritten"
		}
		catch
		{
			Add-content  $Log2 "For $userid $iCity Exception occured"
		}
		
	}
	
	###############################################Country##############################################
	if ($Country -like $null)
	{
		Write-host "$userid has blank Country"
		$Log1 = ".\logs\" + "BlankCountry" + $now + ".log"
		Add-content  $Log1 "$userid has blank Country"
		try
		{
			$Log3 = ".\logs\" + "SetCountry" + $now + ".log"
			Set-ADUser -id $userid -replace @{ co = $iCountry }
			Set-ADUser -id $userid -replace @{ c = $iCountryCode }
			Add-content  $Log3 "For $userid $iCountry as Country has been set"
			Write-Host "For $userid $iCountry as Country has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $iCountry Exception occured"
		}
		
	}
	
	else
	{
		try
		{
			$Log2 = ".\logs\" + "CurrentCountry" + $now + ".log"
			Set-ADUser -id $userid -replace @{ co = $iCountry }
			Set-ADUser -id $userid -replace @{ c = $iCountryCode }
			Add-content  $Log2 "For $userid $iCountry as Current Country has been OverWritten"
			Write-host "$userid has been overwritten with $iCountry as Current Country"
		}
		catch
		{
			Add-content  $Log2 "For $userid $iCountry Exception occured"
		}
		
	}
	
	###############################################Postal Code##########################################################
	
	if ($Zip -like $null)
	{
		Write-host "$userid has blank Zip"
		$Log1 = ".\logs\" + "BlankZip" + $now + ".log"
		Add-content  $Log1 "$userid has blank Zip"
		try
		{
			$Log3 = ".\logs\" + "SetZIP" + $now + ".log"
			Set-ADUser -id $userid -PostalCode $iZip
			Add-content  $Log3 "For $userid $iZip as ZIP has been set"
			Write-Host "For $userid $iZip as ZIP has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $iZip Exception occured"
		}
	}
	
	else
	{
		try
		{
			$Log2 = ".\logs\" + "CurrentZip" + $now + ".log"
			Set-ADUser -id $userid -PostalCode $iZip
			Add-content  $Log2 "For $userid $iZip as Current ZIP has been OverWritten"
			Write-host "$userid has been overwritten with $iZip as Current ZIP"
		}
		catch
		{
			Add-content  $Log2 "For $userid $iZip Exception occured"
		}
	}
	###############################################State################################################################
	
	if ($State -like $null)
	{
		Write-host "$userid has blank State"
		$Log1 = ".\logs\" + "BlankState" + $now + ".log"
		Add-content  $Log1 "$userid has blank State"
		try
		{
			$Log3 = ".\logs\" + "SetState" + $now + ".log"
			Set-ADUser -id $userid -State $istate
			Add-content  $Log3 "For $userid $istate as State has been set"
			Write-Host "For $userid $istate as State has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $istate Exception occured"
		}
		
	}
	
	else
	{
		try
		{
			$Log2 = ".\logs\" + "CurrentState" + $now + ".log"
			Set-ADUser -id $userid -State $istate
			Add-content  $Log2 "For $userid $istate as Current State has been OverWritten"
			Write-host "$userid has been overwritten with $istate as Current State"
		}
		catch
		{
			Add-content  $Log2 "For $userid $istate Exception occured"
		}
	}
	
	###############################################Phone###############################################################
	
	if ($Phone -like $null)
	{
		Write-host "$userid has blank Phone"
		$Log1 = ".\logs\" + "BlankPhone" + $now + ".log"
		Add-content  $Log1 "$userid has blank Phone"
		try
		{
			$Log3 = ".\logs\" + "SetState" + $now + ".log"
			Set-ADUser -id $userid -replace @{ telephoneNumber = $ioffice }
			Add-content  $Log3 "For $userid $ioffice as Telephone has been set"
			Write-Host "For $userid $ioffice as Telephone has been set"
		}
		catch
		{
			Add-content  $Log3 "For $userid $ioffice Exception occured"
		}
		
	}
	
	else
	{
		try
		{
			$Log2 = ".\logs\" + "CurrentPhone" + $now + ".log"
			Set-ADUser -id $userid -replace @{ telephoneNumber = $ioffice }
			Add-content  $Log2 "For $userid $ioffice as Current Telephone has been OverWritten"
			Write-host "$userid has been overwritten with $ioffice as Current Telephone"
		}
		catch
		{
			Add-content  $Log2 "For $userid $ioffice Exception occured"
		}
	}
}
####################################################################################################################

