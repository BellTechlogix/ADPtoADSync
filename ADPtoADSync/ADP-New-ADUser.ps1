#
# ADP_New_ADUser.ps1
#
# Created by Kristopher Roy
# Created Apr 20 2020
# Modified April 20 2020
# Script purpose - Create AD User on Import from ADP

#Source Variables
$sourcedir = "\\BTL-DC-FTP01\"
$sourcefile = "ADPNewUsers.csv"
$ADcreatelog = "ADPtoADNewUsers.log"
$date = get-date -Format "yyyy-MM-dd"

#Required Modules
Import-Module ActiveDirectory

#import the source file for new users from ADP
$userlist = Import-Csv $sourcedir$sourcefile

#write to log
$log = $sourcedir$ADcreatelog
Add-Content $log -Value $date

#loop through each user verify ID doesn't exist then create user
FOREACH($user in $userlist)
{
	$user.name|Add-Content $log
	$aduser = get-aduser -filter 'employeenumber -like $user.associateID'
	If($aduser -eq $null)
	{
		Add-Content $log -Value "Creating New User:"
		Add-Content $log -Value $user.lname", "$userlist.fname
		Add-Content $log -Value "Checking Username Availability:"
	}
}