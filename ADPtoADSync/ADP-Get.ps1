#Rest Method for getting employees directly from ADP
# Created by Kristopher Roy
# Created April 03 2020
# Modified April 03 2020
# Script purpose - Grab ADP details Directly from ADP

$test = Invoke-RestMethod -Uri 'https://test-api.adp.com/hr/v2/workers?$filter=undefined&$select=undefined&$skip=undefined&$top=undefined&$count=undefined' -Method Get
$test.workers.workAssignments
