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
	manager = Employees Supervisor/Manager
	mobile = company cell number
	telephoneNumber = company number
	title = Job Title
#>

#definition for department codes
$deptlookup = @{'710' = "710 - BI Corp Administration";'11002' = "11002 - Epson Depot";'11007' = "11007 - Epson Depot - HR";'14002' = "14002 - Altria Tech Services";'15102' = "15102 - Asset Management Services";
	'15402' = "15402 - HII Services";'20202' = "20202 - TLC Indiana";'20502' = "20502 - TLC Virgina";'21102' = "21102 - USF Services";'30502' = "30502 - Deskside Services";'55002' = "55002 - Service Desk";
	'55005' = "55005 - Service Desk - Mgmt";'60005' = "60005 - Management";'61002' = "61002 - Project Management";'70003' = "70003 - Product - Operations";'74502' = "74502 - Engineering - Tech Ops";
	'75002' = "75002 - Engineering - Projects";'75005' = "75005 - Engineering - Mgmt";'75502' = "75502 - Mobility Solutions";'75505' = "75505 - Mobility Solutions - Mgmt";'77002' = "77002 - Service Delivery Management";
	'79008' = "79008 - Marketing";'79504' = "79504 - Sales - Business Development";'90006' = "90006 - Headquarters - Accounting";'90007' = "90007 - Headquarters - HR";'92509' = "92509 - Headquarters - IT"}

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
    #Split out ADP FirstName and LastName
    $adpln = (($adpuser."employee_name").split(",")[0]).trim()
    $adpfn = ((($adpuser."employee_name").split(",")[1]).trim()).split("")[0].trim()

	#Get ActiveDirectory User from email address
	$email = ($ADPUser."Work Contact: Work Email").trim()
    $aduser = Get-ADUser -Filter{emailaddress -eq $email} -properties *
    
	#check if ADUser is null then try again
    if($aduser -eq $null -or $aduser -eq "")
    {
       #match ADPUser last name
       $aduser = get-aduser -filter{sn -eq $adpln} -Properties *
       #If you get more then one on last name, then match first name
       if($aduser.count -gt 1){$aduser = $aduser|where{$_.givenName -eq $adpfn}}
    }
	#check if ADUser is still null then try again
    if($aduser -eq $null -or $aduser -eq "")
    {
        $adpemail = (($adpfn[0]+$adpln+"@belltechlogix.com").ToLower()).trim()
        $aduser = Get-ADUser -Filter{emailaddress -eq $adpemail} -properties *
    }
	#check if ADUser is still null, if not then proceed, else skip user
    if($aduser -ne $null){
		if($ADPUser."Work Contact: Work Mobile" -and $null -or $ADPUser."Work Contact: Work Mobile" -ne ""){$aduser|Set-ADUser -MobilePhone $ADPUser."Work Contact: Work Mobile"}
		if($ADPUser."Work Contact: Work Phone" -and $null -or $ADPUser."Work Contact: Work Phone" -ne ""){$aduser|Set-ADUser -OfficePhone $ADPUser."Work Contact: Work Phone"}
		if($ADPUser.DeptNumber -ne $null -and $ADPUser.DeptNumber -ne ""){$aduser|Set-ADUser -department $deptlookup[$ADPUser.DeptNumber.trim()]}
		if($ADPUser."Location City" -ne $null -and $ADPUser."Location City" -ne ""){$aduser|Set-ADUser -City $ADPUser."Location City"}
		if($ADPUser."Location Postal Code" -ne $null -and $ADPUser."Location Postal Code" -ne ""){$aduser|Set-ADUser -PostalCode $ADPUser."Location Postal Code"}
		if($ADPUser."Location State/Territory" -ne $null -and $ADPUser."Location State/Territory" -ne ""){$aduser|Set-ADUser -State $ADPUser."Location State/Territory"}
		if($ADPUser."Location Address Line 1" -ne $null -and $ADPUser."Location Address Line 1" -ne ""){$aduser|Set-ADUser -StreetAddress $ADPUser."Location Address Line 1"}
        if($ADPUser."Location Address Line 1" -eq $null -or $ADPUser."Location Address Line 1" -eq ""){$aduser|Set-ADUser -StreetAddress "REMOTE"}
        if($ADPUser."Location Address Line 1" -eq $null -or $ADPUser."Location Address Line 1" -eq ""){$aduser|Set-ADUser -Office "REMOTE"}
        if($ADPUser."Location Address Line 1" -ne $null -and $ADPUser."Location Address Line 1" -ne ""){$aduser|Set-ADUser -Office $ADPUser."Location Description"}
		if($ADPUser.employeeID -ne $null -and $ADPUser.employeeID -ne ""){$aduser|Set-ADUser -EmployeeID $ADPUser.employeeID}
		if($ADPUser.jobtitle -ne $null -and $ADPUser.jobtitle -ne ""){$aduser|Set-ADUser -Title $ADPUser.jobtitle}
		if($ADPUser.jobtitle -ne $null -and $ADPUser.jobtitle -ne ""){$aduser|Where-Object{$_.description -inotlike "*Service*"}|Set-ADUser -Description $ADPUser.jobtitle}
    	$aduser|Where-Object{$_.UserPrincipalName -like "*Service*"}|Set-ADUser -Description ("$ADPfn $adpln Service Account")
		if($ADPUser.SupervisorName -ne $null -and $ADPUser.SupervisorName -ne "")
		{
			$mgr = $ADPUser.SupervisorName
			#split and trim the manager field input to search AD for the user object
			$mgrfn = ($mgr.split(",")[1]).Trim()
			$mgrln = $mgr.split(",")[0].Trim()
			$mgrname = "*$mgrfn*$mgrln*"
			$Manager = Get-ADuser -Filter {Name -like $mgrname}
            #If Manager name doesn't match, try to create a match on an email based on name
			if($Manager -eq $null -or $Manager -eq "")
            {
                $mgremail = (($mgrfn[0]+$mgrln+"@belltechlogix.com").ToLower()).trim()
                #try to match UPN first
				$Manager = Get-ADuser -Filter {userprincipalname -like $mgremail}
            }
			IF($Manager -eq $null -or $Manager -eq "")
            {
                $mgremail = (($mgrfn[0]+$mgrln+"@belltechlogix.com").ToLower()).trim()
                $Manager = Get-ADuser -Filter {emailaddress -like $mgremail}
			}
            $aduser|Set-ADUser -Manager $Manager
		}
	}

	#clear variables from memory so that no accidental write occurs to wrong user
	$aduser = $null
	$ADPuser = $null
    $adpln = $null
    $adpfn = $null
    $email = $null
    $mgr = $null
    $mgrfn = $null
    $mgrln = $null
    $mgrname = $null
	$mgremail = $null
	$adpemail = $null
}
