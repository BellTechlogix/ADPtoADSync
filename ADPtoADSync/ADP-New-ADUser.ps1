#
# ADP_New_ADUser.ps1
#
# Created by Kristopher Roy
# Created Apr 20 2020
# Modified April 20 2020
# Script purpose - Create AD User on Import from ADP

#Source Variables
$sourcedir = "\\BTL-DC-FTP01\c$\FTP\"
$sourcefile = "ADPNewUsers.csv"
$ADcreatelog = "ADPtoADNewUsers.log"
$log = "$sourcedir$ADcreatelog"
$date = get-date -Format "yyyy-MM-dd (%H:mm:ss)"

#Required Modules
Import-Module ActiveDirectory

#import the source file for new users from ADP
$userlist = Import-Csv $sourcedir$sourcefile

#Create log if not exist
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

$date|Add-Content $log

#loop through each user verify ID doesn't exist then create user
FOREACH($user in $userlist)
{
	"   "+$user.name|Add-Content $log
	$aduser = get-aduser -filter 'employeenumber -like $user.associateID'
	If($aduser -eq $null)
	{
		"     Creating New User:"|Add-Content $log
		"     "+$user.lname+", "+$user.fname|Add-Content $log
		"     Checking Username Availability:"|Add-Content $log
        $lnamecount = [Math]::Min($user.lname.Length, 18)
        IF([bool](get-aduser ($user.fname.substring(0,1)+$user.lname.substring(0,$lnamecount))) -eq $true)
        {
            $user.fname.substring(0,1)+$user.lname.substring(0,$lnamecount)+" already exists attempting "|
        }
	}
}


$user = get-aduser kroy|select *,fname,lname
$user.lname = $user.Surname
$user.fname = $user.GivenName