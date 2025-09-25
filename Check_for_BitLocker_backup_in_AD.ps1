# Get the local computer name
$computerName = $env:COMPUTERNAME

# Search AD for recovery keys associated with this computer
Import-Module ActiveDirectory

$searchBase = "DC=distinguished,DC=name"  # Replace with your domain DN

$keys = Get-ADObject -Filter { objectClass -eq "msFVE-RecoveryInformation" } -SearchBase $searchBase -Properties msFVE-RecoveryPassword |
    Where-Object { $_.DistinguishedName -like "*CN=$computerName,*" }

if ($keys) {
    Write-Output "BitLocker recovery keys for $computerName ARE backed up to Active Directory."
} else {
    Write-Output "No BitLocker recovery keys found for $computerName in Active Directory."
}
