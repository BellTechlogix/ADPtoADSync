#
# ADPtoAD-Changes.ps1
# Created by Kristopher Roy
# Created April 24 2020
# Modified April 27 2020
# Script purpose - Write ADP details back to AD Attribute
<#
	AD Attribute Details for use in Script now/or later
	l = location(City)
	postalCode = zipcode
	st = State
	streetAddress = street address
	mail = emailaddress
	employeeID = WhatFromADP
	Department = 
	c = CountryCode
	cn = Name
	name = full name
	co = Country Name
	company = company
	countryCode - 840=US
	department = Dept Accounting Codes
	givenName = FirsName
	sn = LastName
	homePostalAddress = 
	manager = Employees Supervisor/Manager
	mobile = company cell number
	telephoneNumber = company number
	title = Job Title
#>

#Source Variables
$sourcedir = "\\Btl-dc-ftp01\adp\"
$sourcefile = "receive\AD Pull - Users Updated Today.csv"
$date = get-date -Format "yyy-MM-dd"
$timestamp = get-date -Format "yyyy-MM-dd (%H:mm:ss)"
$archivedir = $sourcedir+"archived\UpdatedUsers"
#hrmail recipients for sending report
$hrrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>")
#hdmail recipients for sending report
$hdrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>")
#from address
$from = "BTL-AccountMod@belltechlogix.com"
#smtpserver
$smtp = "smtp.belltechlogix.com"

#definition for department codes
$deptlookup = @{
    '710' = "710 - Corp Administration";
    '720' = "720 - Corp Finance";
    '740' = "740 - Corp Human Resources";
    '11002' = "11002 - Epson Depot";
    '11007' = "11007 - Epson Depot - HR";
    '14002' = "14002 - Altria - TLP";
    '15102' = "15102 - Asset Management Services";
	'15402' = "15402 - HII Services";
    '20202' = "20202 - Indiana Depot";
    '20502' = "20502 - Virgina Depot";
    '21102' = "21102 - USF Services";
    '30502' = "30502 - Deskside Services";
    '55002' = "55002 - Service Desk";
	'55005' = "55005 - Service Desk - Management";
    '55102' = "55102 - Service Desk Operations";
    '55202' = "55202 - Service Improvement";
    '55502' = "55502 - EUS Technology and Automation";
    '60005' = "60005 - Management";
    '61002' = "61002 - Project Management";
    '70003' = "70003 - Product - Operations";
    '74502' = "74502 - Engineering - Tech Ops";
    '75002' = "75002 - Engineering - Projects";
    '75005' = "75005 - Engineering - Mgmt";
    '75502' = "75502 - Mobility Services";
    '75505' = "75505 - Mobility Services Mgmt";
    '77002' = "77002 - Service Delivery Management";
    '79008' = "79008 - Marketing";
    '79504' = "79504 - Sales - Business Development";
    '90006' = "90006 - Headquarters - Accounting";
    '90007' = "90007 - Headquarters - HR";
    '92509' = "92509 - Headquarters - IT"
}

#File Select Function
function Get-FileName
{
  param(
      [Parameter(Mandatory=$false)]
      [string] $Filter,
      [Parameter(Mandatory=$false)]
      [switch]$Obj,
      [Parameter(Mandatory=$False)]
      [string]$Title = "Select A File"
    )
   if(!($Title)) { $Title="Select Input File"}
  
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.FileName = $Title
  #can be set to filter file types
  IF($Filter -ne $null){
  $FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
	$OpenFileDialog.filter = $FilterString}
  if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
  $OpenFileDialog.filter = $Filter
  }
  $OpenFileDialog.ShowDialog() | Out-Null
  IF($OBJ){
  $fileobject = GI -Path $OpenFileDialog.FileName.tostring()
  Return $fileObject
  }
  else{Return $OpenFileDialog.FileName}
}

#ADP data import
# FOR MANUAL Import uncomment this line $ADPFile = Get-FileName -Filter csv -Title "Select ADP Import File"  -Obj
$ADPUsers = Import-Csv $sourcedir$sourcefile

#Create New Error Log File
$log = "$sourcedir\ADP-Modify.log"
    if (!(Test-Path "$log"))
    {
       New-Item -path $sourcedir -type "file" -name $log
       Write-Host "Created new logfile $log"
    }
    else
    {
      Write-Host "Logfile already exists and new text content added"
    }

#Write Timestamp
$timestamp|Add-Content $log
"Updating AD Account info from ADP:"
"_______________________________________"|Add-Content $log

#Loop though users in ADPFile import, match them to AD then write the attributes
FOREACH($User in $ADPUsers)
{
    #Get employee from employee ID
    $ID = $user."Associate ID"
	"   Getting AD account from ADP Associate ID - $ID, User "+$user."Last Name"+", "+$user."First Name:"|Add-Content $log
	$ErrorActionPreference = 'stop'
    try{$aduser = get-aduser -filter 'employeenumber -like $ID' -ErrorAction SilentlyContinue -Properties *}
	catch{"   Unable to match $ID to any AD Accounts"}
	$ErrorActionPreference = 'continue'
	
	#IF ID Matches an AD account update AD Attributes
	IF($aduser -ne $null)
	{
		IF($aduser.employeenumber -eq $ID)
		{
								
			$modifymsg = "Updating field from ADP:"
			
			#Set Variables
			$user."Home Department Code" = $deptlookup[$user."Home Department Code".trim().trimstart('0')]

			#get Managers AD Account
		    $managerID = $user."Reports To Associate ID"
		    $manager = get-aduser -filter 'employeenumber -like $managerID' -properties DistinguishedName -ErrorAction SilentlyContinue

			#Check For Middle Initial and create Name variable
			IF($user.'Middle Initial' -ne $null -and $user.'Middle Initial' -ne "")
				{$Name = ($user."First Name"+" "+$user.'Middle Initial'+" "+$user."Last Name" )}
			ELSE{$Name = ($user."First Name"+" "+$user."Last Name" )}

			#Match and Modify ADAccount info:
			#FirstName
			IF($user."First Name" -ne $aduser.GivenName){"     GivenName "+$aduser.GivenName+" to "+$user."First Name"+" "+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -GivenName $user."First Name"}
			
			#LastName
			IF($user."Last Name" -ne $aduser.Surname){"     Surname "+$aduser.Surname+" to "+$user."Last Name"+" "+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -Surname $user."Last Name"}
			
			#DisplayName
			IF($Name -ne $aduser.displayname){"     Displayname "+$aduser.displayname+" to "+$Name+" "+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -displayname $Name}

			#Name
			IF($Name -ne $aduser.name){"     Name "+$aduser.name+" to "+$Name+" "+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -name $Name}

			#Department
			IF($user."Home Department Code" -ne $aduser.department){"     Department "+$aduser.department+" to "+$user."Home Department Code"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -department $user."Home Department Code"}

			#Manager
			IF($manager.DisguishedName -ne $aduser.Manager){"     Manager "+$aduser.Manager+" to "+$manager.DistinguishedName+" "+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -manager $manager}

			#Title
			IF($user.'Job Title Description' -ne $aduser.Title){"     Title "+$aduser.Title+" to "+$user."Job Title Description"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -Title $user."Job Title Description"}

			#Office
			IF($user.'Location Code' -ne $aduser.Office){"     Office "+$aduser.Office+" to "+$user."Location Code"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -Office $user."Location Code"}

			#Office Address
			IF($user.'Location Description' -ne $aduser.StreetAddress){"     Office Address "+$aduser.StreetAddress+" to "+$user."Location Description"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -StreetAddress $user."Location Description"}

			#Office Phone
			IF($user.'Work Contact: Work Phone' -ne $aduser.OfficePhone){"     Work Phone "+$aduser.OfficePhone+" to "+$user."Work Contact: Work Phone"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -OfficePhone $user."Work Contact: Work Phone"}

			#Mobile Phone
			IF($user.'Personal Contact: Personal Mobile' -ne $aduser.MobilePhone){"     Mobile Phone "+$aduser.MobilePhone+" to "+$user."Personal Contact: Personal Mobile"+$modifymsg|Add-Content $log
			$aduser|Set-ADUser -MobilePhone $user."Personal Contact: Personal Mobile"}
		
			<# Section To Implement once included
				
			If(($user."Location City" -ne $null -and $user."Location City" -ne "") -and $user."Location City" -ne $aduser.city)
			{
				"     City "+$aduser.city+" to "+$user."Location City"+$modifymsg|Add-Content $log
				$aduser|Set-ADUser -City $ADPUser."Location City"
			}
			if(($user."Location Postal Code" -ne $null -and $user."Location Postal Code" -ne "") -and $user."Location Postal Code" -ne $aduser.PostalCode)
			{
				"     Zip "+$aduser.PostalCode+" to "+$user."Location Postal Code"+$modifymsg|Add-Content $log
				$aduser|Set-ADUser -PostalCode $ADPUser."Location Postal Code"
			}
			if(($user."Location State/Territory" -ne $null -and $user."Location State/Territory" -ne "") -and $user."Location State/Territory" -ne $aduser.State)
			{
				"     State "+$aduser.State+" to "+$user."Location State/Territory"+$modifymsg|Add-Content $log
				$aduser|Set-ADUser -State $ADPUser."Location State/Territory"
			}
			#>
		
		}
	}

	#clear variables from memory so that no accidental write occurs to wrong user
	$aduser = $null
	$user = $null
	$name = $null
    $managerID = $null
    $manager = $null
}
