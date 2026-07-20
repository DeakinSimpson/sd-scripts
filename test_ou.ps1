$ou_user_detail = Get-ADUser -Filter "SamAccountName -eq 'deakin.simpson'" |
    Select-Object @{
        Name='OU'
        Expression={
            ($_.DistinguishedName -replace '^CN=.*?,') -replace ',DC=.*$'
        }
    }

$ou_string = $ou_user_detail.OU

$ou_split = $ou_string.split(",")

foreach ($part in $ou_split) {
    $is_disabled = $false

    $prefix_removed = $part -replace "OU="

    $lowercase = $prefix_removed.ToLower()

    Write-Host $lowercase
}