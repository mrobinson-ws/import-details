Connect-AzureAD
Connect-ExchangeOnline

Param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelDoc
)

$users = Import-Excel -Path $ExcelDoc

foreach ($user in $users){
    Set-AzureADUser -ObjectId $user.Email -City $user.City -State $user.State
    Set-Mailbox -Identity $user.email -CustomAttribute1 $user.Type
}