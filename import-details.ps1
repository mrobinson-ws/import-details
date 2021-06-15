[Cmdletbinding()]
Param()

# Test And Connect To AzureAD If Needed
try {
    Write-Verbose -Message "Testing connection to Azure AD"
    Get-AzureAdDomain -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Azure AD"
}
catch {
    Write-Verbose -Message "Connecting to Azure AD"
    Connect-AzureAD
}

#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}
Param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelDoc
)

$users = Import-Excel -Path $ExcelDoc

foreach ($user in $users){
    Set-AzureADUser -ObjectId $user.Email -City $user.City -State $user.State
    Set-Mailbox -Identity $user.email -CustomAttribute1 $user.Type
}