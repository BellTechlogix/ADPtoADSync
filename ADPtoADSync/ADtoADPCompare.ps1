# ADPtoADPCompare.ps1
# Created by Kristopher Roy
# Created Sept 01 2007
# Modified April 03 2020
# Script purpose - compare ADP output details back to AD prior to any writing changes

#csv file output/dump from ADP
$adpinput = import-csv C:\Belltech\ADPOutput_09_22_17.csv

#AD Get
$ADusers = get-aduser -filter * -Properties userprincipalname,department,sn,MobilePhone,OfficePhone,City,postalCode,state,street,employeeID,mail,title,manager,description|select UserPrincipalName,givenName,sn,mail,department,MobilePhone,OfficePhone,City,postalCode,state,street,employeeID,title,manager,description,ADPName,NoADPMatch

#loop Compare AD and ADP
FOREACH($USER in $ADUSers)
{
    $adpuser = $adpinput|where{($_."Work Contact: Work Email").trim() -eq $user.mail}
    If($adpuser -eq $null -or $adpuser -eq "")
    {
        #$adpuser = $adpinput|where{($_."employee_name").split(",")[0] -eq $user.sn}
        #($adpuser.Employee_name.split(",")[1]).trim().split("")[0]
        IF($User.sn -ne $null){
        $adpmatch = $User.sn+", "+$User.givenName
        $adpuser = $adpinput|where{$_.Employee_name -like "*"+$adpmatch+"*"}
        #$adpln = (($adpinput|where{($_."employee_name").split(",")[0] -eq $user.sn})."employee_name").split(",")[0]
        #$adpfn = (($adpinput|where{(($_."employee_name").split(",")[1]).split("")[0] -eq $user.givenName})."employee_name").split(",")[1]
        #-and $_.DeptNumber+"*" -like $user.department}
        #Write-Host $aduser "no match"
        #$adpuser
        $adpmatch = $null
        }
    }
    If($adpuser -ne $null -and $adpuser -ne "")
    {
        $user.ADPName = $adpuser.Employee_name
        $user.NoADPMatch = "False"
        #$user.mail
        #$adpuser.Employee_name
    }
    $adpuser = $null
}

$ADUSers|export-csv c:\belltech\ADP_AD_compare_09_26_17.csv