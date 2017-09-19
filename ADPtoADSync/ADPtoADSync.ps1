#
# ADPtoADSync.ps1
# Created by Kristopher Roy
# Script purpose - Write ADP details back to AD Attribute
<#
	AD Attribute Details for use in Script now/or later
	l = location(City)
	mail = emailaddress
	employeeID = WhatFromADP
	Department = 
	c = CountryCode
	cn = Name
	co = Country Name
	company = company
	countryCode - 840=US
	department = Dept Accounting Codes
	givenName = FirsName
	surName = LastName
	homePostalAddress = 
	manager = has to be full CN for instance (CN=name,OU=whatever,OU=whatever,DC=btl,DC=net)
#>

#definition for department codes
$deptlookup = @{710 = "BI Corp Administration";11002 = "Epson Depot";11007 = "Epson Depot - HR";14002 = "Altria Tech Services";15102 = "Asset Management Services";
	15402 = "HII Services";20202 = "TLC Indiana";20502 = "TLC Virgina";21102 = "USF Services";30502 = "Deskside Services";55002 = "Service Desk";55005 = "Service Desk - Mgmt";
	60005 = "Management";61002 = "Project Management";70003 = "Product - Operations";74502 = "Engineering - Tech Ops";75002 = "Engineering - Projects";75005 = "Engineering - Mgmt";
	75502 = "Mobility Solutions";75505 = "Mobility Solutions - Mgmt";77002 = "Service Delivery Management";79008 = "Marketing";79504 = "Sales - Business Development";
	90006 = "Headquarters - Accounting";90007 = "Headquarters - HR";92509 = "Headquarters - IT"}
