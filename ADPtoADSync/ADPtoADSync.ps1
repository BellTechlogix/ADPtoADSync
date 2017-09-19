#
# ADPtoADSync.ps1
# Created by Kristopher Roy
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
	manager = has to be full CN for instance (CN=name,OU=whatever,OU=whatever,DC=btl,DC=net)
	mobile = company cell number
	telephoneNumber = company number
	title = Job Title
#>

#definition for department codes
$deptlookup = @{710 = "BI Corp Administration";11002 = "Epson Depot";11007 = "Epson Depot - HR";14002 = "Altria Tech Services";15102 = "Asset Management Services";
	15402 = "HII Services";20202 = "TLC Indiana";20502 = "TLC Virgina";21102 = "USF Services";30502 = "Deskside Services";55002 = "Service Desk";55005 = "Service Desk - Mgmt";
	60005 = "Management";61002 = "Project Management";70003 = "Product - Operations";74502 = "Engineering - Tech Ops";75002 = "Engineering - Projects";75005 = "Engineering - Mgmt";
	75502 = "Mobility Solutions";75505 = "Mobility Solutions - Mgmt";77002 = "Service Delivery Management";79008 = "Marketing";79504 = "Sales - Business Development";
	90006 = "Headquarters - Accounting";90007 = "Headquarters - HR";92509 = "Headquarters - IT"}

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
   if(!($Title)) { $Title="Select Input File"} ## why not a default like i showed?
  
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
$ADPFile = Get-FileName -Filter csv -Title "Select ADP Import File" 
$ADPUsers = Import-Csv $ADPFile

#Loop though users in ADPFile import, match them to AD then write the attributes
FOREACH($ADPUser in $ADPUsers)
{
	#Get ActiveDirectory User from email address
	$aduser = get-aduser -Filter{emailaddress -eq $ADPUser.email}
	#check if ADUser is null, if not then proceed, else skip user
	if($aduser -ne $null){
		if($ADPUser.mobile -ne $null -or $ADPUser.mobile -ne ""){$aduser|Set-ADUser -MobilePhone $ADPUser.mobile}
		if($ADPUser.telephone -ne $null -or $ADPUser.telephone -ne ""){$aduser|Set-ADUser -OfficePhone $ADPUser.telephone}
		if($ADPUser.deptcode -ne $null -or $ADPUser.deptcode -ne ""){$aduser|Set-ADUser -department $deptlookup[$ADPUser.deptcode]}
		}

	#clear variables from memory so that no accidental write occurs to wrong user
	$aduser = $null
	$ADPuser = $null
}
