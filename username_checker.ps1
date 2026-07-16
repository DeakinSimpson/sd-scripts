$Usernames = Import-Csv -Path "C:\temp\usernames.csv"
 
foreach ($Username in $Usernames) {
$User = $Username.Usernames
if (Get-ADUser -Filter "SamAccountName -eq '$User'") {
    Write-Output "User '$User' exists in Active Directory."
} else {
    Write-Output "User '$User' does NOT exist."
}
}