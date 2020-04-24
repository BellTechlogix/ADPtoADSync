#
# ADPtoAD-Changes.ps1
# Created by Kristopher Roy
# Created April 24 2020
# Modified April 24 2020
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
$sourcedir = "\\BTL-DC-FTP01\c$\FTP\"
$sourcefile = "AD Pull - Users Updated Today.csv"
$date = get-date -Format "yyy-MM-dd"
$timestamp = get-date -Format "yyyy-MM-dd (%H:mm:ss)"
$archivedir = $sourcedir+"archived\UpdatedUsers"
#hrmail recipients for sending report
$hrrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>")
#hdmail recipients for sending report
$hdrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>")
#from address
$from = "BTL-AccountCreation@belltechlogix.com"
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
# FOR MANUAL Import $ADPFile = Get-FileName -Filter csv -Title "Select ADP Import File"  -Obj
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
    try{$aduser = get-aduser -filter 'employeenumber -like $ID' -ErrorAction SilentlyContinue -Properties employeenumber}
	catch{"   Unable to match $ID to any AD Accounts"}
	$ErrorActionPreference = 'continue'
	
	#IF ID Matches an AD account update AD Attributes
	IF($aduser -ne $null)
	{
		IF($aduser.)
	}
	
	
    
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
		if($ADPUser."Work Contact: Work Mobile" -ne $null -and $ADPUser."Work Contact: Work Mobile" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Mobile Number Updated -Original:"+$aduser.MobilePhone+" -New:"+$ADPUser."Work Contact: Work Mobile"
			$aduser|Set-ADUser -MobilePhone $ADPUser."Work Contact: Work Mobile"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser."Work Contact: Work Phone" -ne $null -and $ADPUser."Work Contact: Work Phone" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Office Number Updated -Original:"+$aduser.OfficePhone+" -New:"+$ADPUser."Work Contact: Work Phone"
			$aduser|Set-ADUser -OfficePhone $ADPUser."Work Contact: Work Phone"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser.DeptNumber -ne $null -and $ADPUser.DeptNumber -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Department Number Updated -Original:"+$aduser.department+" -New:"+$deptlookup[$ADPUser.DeptNumber.trim()]			
			$aduser|Set-ADUser -department $deptlookup[$ADPUser.DeptNumber.trim()]
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null		
		}
        if($ADPUser."Location Address Line 1" -eq $null -or $ADPUser."Location Address Line 1" -eq "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Street Address Updated -Original:"+$aduser.StreetAddress+" -New:REMOTE"			
			$aduser|Set-ADUser -StreetAddress "REMOTE"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Office Updated -Original:"+$aduser.Office+" -New:REMOTE"
			$aduser|Set-ADUser -Office "REMOTE"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - City Updated -Original:"+$aduser.City+" -New:NULL"
			$aduser|Set-ADUser -City $null
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Zip Updated -Original:"+$aduser.PostalCode+" -New:NULL"
			$aduser|Set-ADUser -PostalCode $null
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - State Updated -Original:"+$aduser.State+" -New:NULL"
			$aduser|Set-ADUser -State $null
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
        if($ADPUser."Location Address Line 1" -ne $null -and $ADPUser."Location Address Line 1" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Office Updated -Original:"+$aduser.Office+" -New:"+$ADPUser."Location Description"
			$aduser|Set-ADUser -Office $ADPUser."Location Description"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Street Address Updated -Original:"+$aduser.StreetAddress+" -New:"+$ADPUser."Location Address Line 1"
			$aduser|Set-ADUser -StreetAddress $ADPUser."Location Address Line 1"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser."Location City" -ne $null -and $ADPUser."Location City" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - City Updated -Original:"+$aduser.City+" -New:"+$ADPUser."Location City"
			$aduser|Set-ADUser -City $ADPUser."Location City"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser."Location Postal Code" -ne $null -and $ADPUser."Location Postal Code" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Zip Updated -Original:"+$aduser.PostalCode+" -New:"+$ADPUser."Location Postal Code"
			$aduser|Set-ADUser -PostalCode $ADPUser."Location Postal Code"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser."Location State/Territory" -ne $null -and $ADPUser."Location State/Territory" -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - State Updated -Original:"+$aduser.State+" -New:"+$ADPUser."Location State/Territory"
			$aduser|Set-ADUser -State $ADPUser."Location State/Territory"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser.employeeID -ne $null -and $ADPUser.employeeID -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - State Updated -Original:"+$aduser.EmployeeID+" -New:"+$ADPUser.employeeID
			$aduser|Set-ADUser -EmployeeID $ADPUser.employeeID
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		if($ADPUser.jobtitle -ne $null -and $ADPUser.jobtitle -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - JobTitle Updated -Original:"+$aduser.Title+" -New:"+$ADPUser.jobtitle
			$aduser|Set-ADUser -Title $ADPUser.jobtitle
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Description Updated -Original:"+$aduser.Description+" -New:"+$ADPUser.jobtitle
			$aduser|Where-Object{$_.description -inotlike "*Service*"}|Set-ADUser -Description $ADPUser.jobtitle
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null		
		}
    	$aduser|Where-Object{$_.UserPrincipalName -like "*Service*"}|Set-ADUser -Description ("$ADPfn $adpln Service Account")
		IF(($aduser|Where-Object{$_.UserPrincipalName -like "*Service*"}).UserPrincipalName -like "*Service*"){
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" Service Account - Description Updated -Original:"+$aduser.Description+" -New:$ADPfn $adpln Service Account"
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
		}
		$timestamp = $null
		$logline = $null
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
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" - Manager Updated -Original:"+($aduser.Manager|out-string).split(",")[0].substring(3)+" -New:"+$Manager.name
            $aduser|Set-ADUser -Manager $Manager
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
		#Remove manager for service accounts so they don't show up in the GAL
		$svcaccount = $aduser|Where-Object{$_.UserPrincipalName -like "*Service*"}
		IF($svcaccount -ne $null -and $svcaccount -ne "")
		{
			$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
			$logline = "`n"+$timestamp+" - Success: "+$adpname+" Service Account - Manager Removed -Original:"+($aduser.Manager|out-string).split(",")[0].substring(3)+" -New:Null"
			$svcaccount|Set-ADUser -Manager $null
			Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $logline
			$timestamp = $null
			$logline = $null
		}
	}

	#IF ADuser is still null add to log
	if($aduser -eq $null -or $aduser -eq "")
	{
		$timestamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
		$errorline = "`n"+$timestamp+" Error "+$ADPUser.Employee_name+" No Match Found, AD User Data Not Updated"
		Add-Content ($ADPFile.PSParentPath+"\adpimport.log") $errorline
		$errorline = $null
	}

	#clear variables from memory so that no accidental write occurs to wrong user
	$aduser = $null
	$ADPuser = $null
    $adpln = $null
    $adpfn = $null
	$adpname = $null
    $email = $null
    $mgr = $null
    $mgrfn = $null
    $mgrln = $null
    $mgrname = $null
	$mgremail = $null
	$adpemail = $null
	$svcaccount = $null
}
