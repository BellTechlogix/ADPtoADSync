#
# ADP_New_ADUser.ps1
#
# Created by Kristopher Roy
# Created Apr 20 2020
# Modified April 29 2020
# Script purpose - Create AD User on Import from ADP

#Source Variables
$sourcedir = "\\Btl-dc-ftp01\adp\"
$sourcefile = "receive\AD_Pull-Users_Created_Today.csv"
$date = get-date -Format "yyy-MM-dd"
$timestamp = get-date -Format "yyyy-MM-dd (%H:mm:ss)"
$archivedir = $sourcedir+"archive"
#hrmail recipients for sending report
$hrrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>","BTLHR <HumanResources@belltechlogix.com>")
#hdmail recipients for sending report
$hdrecipients = @("Kristopher <kroy@belltechlogix.com>","Jack <hchen@belltechlogix.com>","HelpDesk <Helpdesk@belltechlogix.com>")
#from address
$from = "BTL-AccountCreation@belltechlogix.com"
#smtpserver
$smtp = "smtp.belltechlogix.com"
#Exchange Connect Session
$remoteex = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://BTL-CORP-CAS01/PowerShell/

#Required Modules
Import-Module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

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

#definition for Mailbox Databases
$mbdblookup = @{
    '710' = "EUC_2";
    '720' = "EXEC_2";
    '740' = "EXEC_2";
    '11002' = "EUC_2";
    '11007' = "EXEC_2";
    '14002' = "EUC_2";
    '15102' = "EUCSUP_2";
    '15402' = "EUC_2";
    '20202' = "STAFF_2";
    '21102' = "STAFF_2";
    '30502' = "EUCSUP_2";
    '55002' = "EUC_2";
    '55005' = "EUCSUP_2";
    '55102' = "EUCSUP_2";
    '55202' = "STAFF_2";
    '55502' = "STAFF_2";
    '60005' = "EXEC_2";
    '61002' = "STAFF_2";
    '70003' = "EXEC_2";
    '74502' = "IMS_2";
    '75002' = "IMS_2";
    '75005' = "EXEC_2";
    '75502' = "EUC_2";
    '75505' = "EUCSUP_2";
    '77002' = "EUCSUP_2";
    '79008' = "STAFF_2";
    '79504' = "EXEC_2";
    '90006' = "STAFF_2";
    '90007' = "EXEC_2";
    '92509' = "IMS_2"
}

#import the source file for new users from ADP
TRY{$userlist = Import-Csv $sourcedir$sourcefile|select *}CATCH{$userlist = $null}
IF($userlist -ne $null)
{
    #for testing
    #$user = $userlist|where{$_."Associate ID" -eq "2TPN84ARF77"}

    #loop through each user verify ID doesn't exist then create user
    FOREACH($user in $userlist)
    {  
        #create userlog
        $ADcreatelog = $user."First Name"+"."+$user."Last Name"+(get-date -Format "yyyy-MM-dd-%H-mm-ss")+".log"
        $log = "$sourcedir$ADcreatelog"
        if (!(Test-Path "$log"))
        {
           TRY{New-Item -path $sourcedir -type "file" -name $ADcreatelog}CATCH{New-Item -path $sourcedir -type "file" -name "unknownuser"+(get-date -Format "yyyy-MM-dd-%H-mm-ss")+".log"}
           Write-Host "Created new logfile $ADcreatelog"
        }
        else
        {
          Write-Host "Logfile already exists and new text content added"
        }


        #write to log
        $timestamp|Add-Content $log
        "   "+$user."Payroll Name"|Add-Content $log
        "   "+$user."Associate ID"|Add-Content $log
        "   "+$user."First Name"+" "+$user."Last Name"|Add-Content $log

	    #set ID
	    $ID = $user."Associate ID"
		#Check For Middle Initial and create Name variable
		IF($user.'Middle Initial' -ne $null -and $user.'Middle Initial' -ne "")
			{$Name = ($user."First Name"+" "+$user.'Middle Initial'+" "+$user."Last Name" )}
		ELSE{$Name = ($user."First Name"+" "+$user."Last Name" )}

	    #Check for missing fields
	    $missingmsg = "missing from source ADP, halting creation and mailing HR"
	    IF($user."First Name" -eq $null -or $user."First Name" -eq ""){"   First Name $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode = "Missing Firstname "}
	    IF($user."Last Name" -eq $null -or $user."Last Name" -eq ""){"   Last Name $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode += "Missing Last Name "
	    }
	    IF($user."Associate ID" -eq $null -or $user."Associate ID" -eq ""){"   Associate ID $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode += "Missing Associate ID "
	    }
        IF($user."Home Department Code" -eq $null -or $user."Home Department Code" -eq ""){"   Home Department Code $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode += "Missing Home Department Code "
	    }
        IF($user.'Job Title Description' -eq $null -or $user.'Job Title Description' -eq ""){"   Job Title Description $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode += "Job Title Description "
	    }
        IF($user.'Location Code' -eq $null -or $user.'Location Code' -eq ""){"   Location Code $missingmsg"|Add-Content $log
		    $missing = 'True'
		    $failcode += "Location Code "
	    }
    
	    If($missing -ne 'True')
	    {
		    #get Managers AD Account
		    $managerID = $user."Reports To Associate ID"
		    $manager = get-aduser -filter 'employeenumber -like $managerID' -ErrorAction SilentlyContinue
    
	
		    #Check that employee ID doesn't exist then create user account
		    $aduser = get-aduser -filter 'employeenumber -like $ID' -ErrorAction SilentlyContinue -Properties employeenumber,whencreated
		    If($aduser -eq $null)
		    {
			    "     Creating New User:"|Add-Content $log
			    "        "+$user."Last Name"+", "+$user."First Name"|Add-Content $log
			    $lnamecount = [Math]::Min($user."Last Name".Length, 18)
			    $initusername = ($user."First Name".substring(0,1)+$user."Last Name".substring(0,$lnamecount))
			    "        Checking Username Availability:"+$initusername|Add-Content $log
			    IF([bool](get-aduser -Filter{SamAccountName -eq $initusername} -ErrorAction SilentlyContinue) -eq $true)
			    {
				    IF($lnamecount -gt 17){$lnamecount = $lnamecount - 1}
				    $ErrorActionPreference = 'stop'
				    try{$secondusername = ($user."First Name".substring(0,1)+$user."Middle Initial".substring(0,1)+$user."Last Name".substring(0,$lnamecount))}
				    catch{$secondusername = ($user."First Name".substring(0,1)+$user."First Name".substring(1,2)+$user."Last Name".substring(0,$lnamecount))}
				    $ErrorActionPreference = 'continue'
				    "        "+$initusername+" already exists attempting "+$secondusername|Add-Content $log
            
				    #check if secondary user name attempt already exists
				    IF([bool](get-aduser -Filter{SamAccountName -eq $secondusername} -ErrorAction SilentlyContinue) -eq $true)
				    {
					    #Send email to helpdesk for no available account names
					    $htmlforHDFailEmail = "<h1 style='color: #5e9ca0;'><span style='text-decoration: underline;'>User Account Creation Fail</span></h1>"
					    $htmlforHDFailEmail = $htmlforHDFailEmail + "<h2 style='color: #2e6c80;'>User Account Names Unavailable for Employee: <span style='color: #000000;'>$ID</span></h2>"
					    $htmlforHDFailEmail = $htmlforHDFailEmail + "<h2 style='color: #2e6c80;'>ADPUSER:&nbsp;<span style='color: #000000;'>"+$USER."First Name"+" "+$USER."Last Name"+"</span></h2>"
					    $htmlforHDFailEmail = $htmlforHDFailEmail + "<h2 style='color: #2e6c80;'>ADUSER Accounts tried:&nbsp;<span style='color: #000000;'>"+$initusername+", "+$secondusername+"</span></h2>"
					    $htmlforHDFailEmail = $htmlforHDFailEmail +  "<h4><span style='color: #000000;'>Please find available user ID then create account and mailbox</span></h4>"
					    "        Usernames:"+$initusername+", "+$secondusername+" already exist, forwarding to ServiceDesk@belltechlogix.com"|Add-Content $log
					    Send-MailMessage -from $from -to $hdrecipients -subject "BTL No Available UserID for Auto-Account Creation" -smtpserver $smtp -BodyAsHtml $htmlforHDFailEmail -Attachments $log
            
				    }
				    ELSEIF([bool](get-aduser -Filter{SamAccountName -eq $secondusername} -ErrorAction SilentlyContinue) -eq $false)
				    {
					    "        Username:"+$secondusername+ " Available, creating account:"|Add-Content $log
					    
						#creating account based upon first initial+middle initial+last name
					    New-ADUser -SamAccountName $secondusername `
					    -UserPrincipalName ($secondusername+"@belltechlogix.com") `
					    -Name $Name `
					    -DisplayName $Name `
					    -Surname $user."Last Name" -GivenName $user."First Name" `
					    -Initials $user.'Middle Initial' -EmployeeNumber $user."Associate ID" `
					    -Department ($deptlookup[$user."Home Department Code".trim().trimstart('0')]) `
					    -Manager $manager -Title $user.'Job Title Description' `
					    -Office $user.'Location Code' -StreetAddress $user.'Location Description' `
					    -OfficePhone $user.'Work Contact: Work Phone' `
					    -MobilePhone $user.'Personal Contact: Personal Mobile' `
                        -City $user.'Location City' `
                        -PostalCode $user.'Location Postal Code' `
                        -State $user.'Location State/Territory' `
					    -path "OU=\#\#Automation_Purgatory,DC=btl,DC=bellind,DC=net" `
					    -Enabled 1 -PasswordNotRequired 1 `
					    -ErrorAction Continue|Add-Content $log			
                
					    Add-Content $log -Value "        account created $secondusername"
					    Add-Content $log -Value "        waiting 30s before creating mailbox $secondusername"
					    Start-Sleep -Seconds 30
                                
					    #try and create mailbox
					    Add-Content $log -Value "        creating mailbox for $secondusername"
					    $ErrorActionPreference = 'stop'        
    				    try{Enable-Mailbox -Identity $secondusername -Database ($mbdblookup[$user."Home Department Code".trim().trimstart('0')])}
					    catch{Invoke-Command -Session $remoteex -ScriptBlock{Enable-Mailbox -Identity $args[0] -Database $args[1]} -ArgumentList $secondusername,($mbdblookup[$user."Home Department Code".trim().trimstart('0')])}
					    $ErrorActionPreference = 'continue'

					    #Send email to helpdesk for succesful account creation with secondary username
					    $htmlforHDsecondsuccessEmail = "<h1 style='color: #5e9ca0;'><span style='text-decoration: underline;'>User Account Creation Success</span></h1>"
					    $htmlforHDsecondsuccessEmail = $htmlforHDsecondsuccessEmail + "<h2 style='color: #2e6c80;'>User Account Created for Employee: <span style='color: #000000;'>$ID</span></h2>"
					    $htmlforHDsecondsuccessEmail = $htmlforHDsecondsuccessEmail + "<h2 style='color: #2e6c80;'>ADPUSER:&nbsp;<span style='color: #000000;'>"+$USER."First Name"+" "+$USER."Last Name"+"</span></h2>"
					    $htmlforHDsecondsuccessEmail = $htmlforHDsecondsuccessEmail + "<h2 style='color: #2e6c80;'>ADUSER Account Created:&nbsp;<span style='color: #000000;'>"+$secondusername+"</span></h2>"
					    $htmlforHDsecondsuccessEmail = $htmlforHDsecondsuccessEmail +  "<h4><span style='color: #000000;'>Please verify account and mailbox success and accuracy</span></h4>"
					    "        Usernames:"+$secondusername+" was succesfully created, forwarding to ServiceDesk@belltechlogix.com for review"|Add-Content $log
					    Send-MailMessage -from $from -to $hdrecipients -subject "BTL Succesfull Auto-Account Creation" -smtpserver $smtp -BodyAsHtml $htmlforHDsecondsuccessEmail -Attachments $log|Add-Content $log
                
					    #clear html 
					    $htmlforHDsecondsuccessEmail = $null
            
				    }
			    }
			    ELSEIF([bool](get-aduser -Filter{SamAccountName -eq $initusername} -ErrorAction SilentlyContinue) -eq $false)
			    {
				    "        Username:"+$initusername+ " Available, creating account:"|Add-Content $log
            
				    #creating account based upon first initial+last name
				    New-ADUser -SamAccountName $initusername `
				    -UserPrincipalName ($initusername+"@belltechlogix.com") `
				    -Name $Name  `
				    -DisplayName $Name `
				    -Surname $user."Last Name" `
				    -GivenName $user."First Name" `
				    -Initials $user.'Middle Initial' `
				    -EmployeeNumber $user."Associate ID" `
				    -Department ($deptlookup[$user."Home Department Code".trim().trimstart('0')]) `
				    -Manager $manager `
				    -Title $user.'Job Title Description' `
				    -Office $user.'Location Code' `
				    -StreetAddress $user.'Location Description' `
				    -OfficePhone $user.'Work Contact: Work Phone' `
				    -MobilePhone $user.'Personal Contact: Personal Mobile' `
				    -path "OU=\#\#Automation_Purgatory,DC=btl,DC=bellind,DC=net" `
				    -Enabled 1 -PasswordNotRequired 1 `
				    -ErrorAction Continue|Add-Content $log

				    Add-Content $log -Value "        account created $initusername"			
				    Start-Sleep -Seconds 30
			
				    #try and create mailbox
				    Add-Content $log -Value "        creating mailbox for $initusername"
				    $ErrorActionPreference = 'stop'        
    			    try{Enable-Mailbox -Identity $initusername -Database ($mbdblookup[$user."Home Department Code".trim().trimstart('0')])}
				    catch{Invoke-Command -Session $remoteex -ScriptBlock{Enable-Mailbox -Identity $args[0] -Database $args[1]} -ArgumentList $initusername,($mbdblookup[$user."Home Department Code".trim().trimstart('0')])}
				    $ErrorActionPreference = 'continue'

				    #Send email to helpdesk for succesful account creation with initial username
				    $htmlforHDInitialsuccessEmail = "<h1 style='color: #5e9ca0;'><span style='text-decoration: underline;'>User Account Creation Success</span></h1>"
				    $htmlforHDInitialsuccessEmail = $htmlforHDInitialsuccessEmail + "<h2 style='color: #2e6c80;'>User Account Created for Employee: <span style='color: #000000;'>$ID</span></h2>"
				    $htmlforHDInitialsuccessEmail = $htmlforHDInitialsuccessEmail + "<h2 style='color: #2e6c80;'>ADPUSER:&nbsp;<span style='color: #000000;'>"+$USER."First Name"+" "+$USER."Last Name"+"</span></h2>"
				    $htmlforHDInitialsuccessEmail = $htmlforHDInitialsuccessEmail + "<h2 style='color: #2e6c80;'>ADUSER Account Created:&nbsp;<span style='color: #000000;'>"+$initusername+"</span></h2>"
				    $htmlforHDInitialsuccessEmail = $htmlforHDInitialsuccessEmail +  "<h4><span style='color: #000000;'>Please verify account and mailbox success and accuracy</span></h4>"
				    "        Usernames:"+$HDInitialusername+" was succesfully created, forwarding to ServiceDesk@belltechlogix.com for review"|Add-Content $log
				    Send-MailMessage -from $from -to $hdrecipients -subject "BTL Succesfull Auto-Account Creation" -smtpserver $smtp -BodyAsHtml $htmlforHDInitialsuccessEmail -Attachments $log|Add-Content $log
                
				    #clear html 
				    $htmlforHDInitialsuccessEmail = $null
			    }
		    }
		    ELSEIF($aduser -ne $null)
		    {
                IF(!($aduser.whencreated -gt (get-date).AddDays(-1))){
			    #Send email to HR for User ID duplicate
			    $htmlforHREmail = "<h1 style='color: #5e9ca0;'><span style='text-decoration: underline;'>User Account Creation Fail</span></h1>"
			    $htmlforHREmail = $htmlforHREmail + "<h2 style='color: #2e6c80;'>Duplicate Employee Number: <span style='color: #000000;'>$ID</span></h2>"
			    $htmlforHREmail = $htmlforHREmail + "<h2 style='color: #2e6c80;'>ADUSER:&nbsp;<span style='color: #000000;'>"+$ADUSER.SamAccountName+"</span></h2>"
			    $htmlforHREmail = $htmlforHREmail + "<h2 style='color: #2e6c80;'>ADPUSER:&nbsp;<span style='color: #000000;'>"+$USER."First Name"+" "+$USER."Last Name"+"</span></h2>"
			    $htmlforHREmail = $htmlforHREmail + "<h4><span style='color: #000000;'>Please Resolve User ID duplicate and contact the helpdesk for account creation</span></h4>"

			    "        Username:"+$aduser.SamAccountName+" with employeID:"+$aduser.employeenumber+" already exists, forwarding to HumanResources@belltechlogix.com"|add-content $log
			    Send-MailMessage -from $from -to $hrrecipients -subject "BTL Pre-Existing employee ID" -smtpserver $smtp -BodyAsHtml $htmlforHREmail -Attachments $log|Add-Content $log
			    $htmlforHREmail = $null
                }
                ELSEIF(!($aduser.whencreated -lt (get-date).AddDays(-1))){Remove-Item $log}
		    }
    
	    }
	    ElseIf($missing -eq 'True')
	    {
	        
		    #Send email to HR for required ADP field missing
		    $htmlforHREmail = "<h1 style='color: #5e9ca0;'><span style='text-decoration: underline;'>User Account Creation Fail</span></h1>"
		    $htmlforHREmail = $htmlforHREmail + "<h2 style='color: #2e6c80;'>ADP Field Missing For: <span style='color: #000000;'>$ID</span></h2>"
		    $htmlforHREmail = $htmlforHREmail + "<h2 style='color: #2e6c80;'>ADPUSER:&nbsp;<span style='color: #000000;'>"+$USER."First Name"+" "+$USER."Last Name"+"</span></h2>"
		    $htmlforHREmail = $htmlforHREmail + "<h4><span style='color: #000000;'>Please Resolve $failcode</span></h4>"

		    "        $failcode, forwarding to HumanResources@belltechlogix.com"|add-content $log
		    Send-MailMessage -from $from -to $hrrecipients -subject "BTL ADP Fields Missing" -smtpserver $smtp -BodyAsHtml $htmlforHREmail -Attachments $log|Add-Content $log
		    $htmlforHREmail = $null
		    $failcode = $null
	    }
		    #Move User log file
		    try{move-item $log -Destination "$archivedir\NewUsers"}CATCH{}

            #clear variables for loop iteration
            $missing = $null
            $failcode = $null
		    $ADcreatelog = $null
		    $log = $null
		    $user = $null
		    $aduser = $null
		    $manager = $null
		    $managerID = $null
			$Name = $null
            Start-Sleep -Seconds 2
    }
    #move source file to archive
    $archivefilename = "ADP-NewUsers-"+(get-date -Format "yyyy-MM-dd-%H-mm-ss")+".csv"
    Rename-Item $sourcedir$sourcefile -NewName $archivefilename
    Move-Item $sourcedir"receive\"$archivefilename -Destination $archivedir
}
$userlist = $null