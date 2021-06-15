[Cmdletbinding()]
Param(
)

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

$users = Import-Excel -Path "C:\Git Repo\import-details\VelocityUserList 6-14.xlsx"

foreach ($user in $users){
    Write-Verbose "Setting $($user.Email) City and State"
    Set-AzureADUser -ObjectId $user.Email -City $user.City -State $user.State
    Write-Verbose "Setting $($user.Email) Custom Attribute 1"
    Set-Mailbox -Identity $user.email -CustomAttribute1 $user.Type
}