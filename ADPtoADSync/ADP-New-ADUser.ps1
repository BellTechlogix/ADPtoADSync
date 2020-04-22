#
# ADP_New_ADUser.ps1
#
# Created by Kristopher Roy
# Created Apr 20 2020
# Modified April 20 2020
# Script purpose - Create AD User on Import from ADP

#Source Variables
$sourcedir = "\\BTL-DC-FTP01\c$\FTP\"
$sourcefile = "AD Pull - Users Created Today (2).csv"
$date = get-date -Format "yyy-MM-dd"
$timestamp = get-date -Format "yyyy-MM-dd (%H:mm:ss)"

#Required Modules
Import-Module ActiveDirectory

#definition for department codes
$deptlookup = @{'710' = "710 - Corp Administration";'720' = "720 - Corp Finance";'740' = "740 - Corp Human Resources";'11002' = "11002 - Epson Depot";'11007' = "11007 - Epson Depot - HR";'14002' = "14002 - Altria - TLP";'15102' = "15102 - Asset Management Services";
	'15402' = "15402 - HII Services";'20202' = "20202 - Indiana Depot";'20502' = "20502 - Virgina Depot";'21102' = "21102 - USF Services";'30502' = "30502 - Deskside Services";'55002' = "55002 - Service Desk";
	'55005' = "55005 - Service Desk - Management";'55102' = "55102 - Service Desk Operations";'55202' = "55202 - Service Improvement";'55502' = "55502 - EUS Technology and Automation";'60005' = "60005 - Management";'61002' = "61002 - Project Management";
    '70003' = "70003 - Product - Operations";'74502' = "74502 - Engineering - Tech Ops";'75002' = "75002 - Engineering - Projects";'75005' = "75005 - Engineering - Mgmt";'75502' = "75502 - Mobility Services";
    '75505' = "75505 - Mobility Services Mgmt";'77002' = "77002 - Service Delivery Management";'79008' = "79008 - Marketing";'79504' = "79504 - Sales - Business Development";'90006' = "90006 - Headquarters - Accounting";
    '90007' = "90007 - Headquarters - HR";'92509' = "92509 - Headquarters - IT"}

#import the source file for new users from ADP
$userlist = Import-Csv $sourcedir$sourcefile|select *,adpln,adpfn,adpMn


#loop through each user verify ID doesn't exist then create user
FOREACH($user in $userlist)
{
	#Split out ADP FirstName, LastName, MiddleName
    $user.adpln = (($user."Payroll Name").split(",")[0]).trim()
    $user.adpfn = ((($user."Payroll Name").split(",")[1]).trim()).split("")[0].trim()
    $user.adpMn = ((($user."Payroll Name").split(" ")[2]).trim()).split("")[0].trim()
    
    #create userlog
    $ADcreatelog = $user.adpfn+"."+$user.adpln+".log"
    $log = "$sourcedir$ADcreatelog"
    if (!(Test-Path "$log"))
    {
       New-Item -path $sourcedir -type "file" -name $ADcreatelog
       Write-Host "Created new logfile $ADcreatelog"
    }
    else
    {
      Write-Host "Logfile already exists and new text content added"
    }


    #write to log
    $timestamp|Add-Content $log
    "   "+$user."Payroll Name"|Add-Content $log
    
	
    #Check that employee ID doesn't exist then create user account
    $ID = $user."Associate ID"
    $aduser = get-aduser -filter 'employeenumber -like $ID' -ErrorAction SilentlyContinue
	If($aduser -eq $null)
	{
		"     Creating New User:"|Add-Content $log
		"        "+$user.adpln+", "+$user.adpfn|Add-Content $log
        $lnamecount = [Math]::Min($user.adpln.Length, 18)
        $initusername = ($user.adpfn.substring(0,1)+$user.adpln.substring(0,$lnamecount))
        "        Checking Username Availability:"+$initusername|Add-Content $log
        IF([bool](get-aduser -Filter{SamAccountName -eq $initusername} -ErrorAction SilentlyContinue) -eq $true)
        {
            IF($lnamecount -gt 17){$lnamecount = $lnamecount - 1}
            $secondusername = ($user.adpfn.substring(0,1)+$user.adpMn.substring(0,1)+$user.adpln.substring(0,$lnamecount))
            "        "+$initusername+" already exists attempting "+$secondusername|Add-Content $log
            IF([bool](get-aduser -Filter{SamAccountName -eq $secondusername} -ErrorAction SilentlyContinue) -eq $true)
            {}
            ELSEIF([bool](get-aduser -Filter{SamAccountName -eq $secondusername} -ErrorAction SilentlyContinue) -eq $false)
            {
                "        Username:"+$secondusername+ " Available, creating account:"|Add-Content $log
                #creating account based upon first initial+middle initial+last name
                New-ADUser -SamAccountName $secondusername -Name ($user.adpfn+" "+$user.adpln ) -Surname $user.adpln -GivenName $user.adpfn -EmployeeNumber $user."Associate ID" -Department ($deptlookup[$user."Home Department Code".trim().trimstart('0')]) -WhatIf|out-file $log -Append
				Start-Sleep -Seconds 30
            }
        }
        ELSEIF([bool](get-aduser -Filter{SamAccountName -eq $initusername} -ErrorAction SilentlyContinue) -eq $false)
        {
            "        Username:"+$initusername+ " Available, creating account:"|Add-Content $log
            #creating account based upon first initial+last name
            New-ADUser -SamAccountName $initusername -Surname $user.adpln -GivenName $user.adpfn -EmployeeNumber $user."Associate ID" -WhatIf|Add-Content $log
			Start-Sleep -Seconds 30
        }
	}
    ELSEIF($aduser -ne $null)
    {
        
    }
    #clear variables for loop iteration
    $ADcreatelog = $null
    $log = $null
    $user = $null
}