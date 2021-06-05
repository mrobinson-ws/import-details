#Require -Modules AzureAD, ImportExcel
Param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelDoc
)

$users = Import-Excel -Path $ExcelDoc

foreach ($user in $users){Set-AzureADUser -ObjectId $user.Email -CustomAttribute1 $user.Type -City $user.City -State $user.State}