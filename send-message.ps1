# variables for api key
# grab the api key from liquidfiles
$apikey = ""

# credential encoding
$pair = "$($apikey):$()"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))

# import the csv
$csvPath = Join-Path -Path (Get-Location) -ChildPath "users.csv"
$users = Import-Csv -Path $csvPath

foreach ($user in $users) {
    $response = $null

    $email = $null
    $ADusername = $null
    $ADpassword = $null
    $IPMusername = $null
    $IPMpassword = $null
    $EMRusername = $null
    $EMRpassword = $null
    $BOSSNETusername = $null
    $BOSSNETpassword = $null
    $ServiceRequest = $null
    $messageText = ""

    $email = $user.'requesters email'
    $ADusername = $user.'AD username'
    $ADpassword = $user.'AD password'
    $IPMusername = $user.'IPM username'
    $IPMpassword = $user.'IPM password'
    $EMRusername = $user.'EMR username'
    $EMRpassword = $user.'EMR password'
    $BOSSNETusername = $user.'BOSSNET username'
    $BOSSNETpassword = $user.'BOSSNET password'
    $ServiceRequest = $user.'SR'

    if (-not $email -or $email.Trim() -eq "") {
        Write-Host "Skipping row with no email"
        continue
    }
    

    if ($ADusername) {
        $messageText += "Computer Account`nUsername: $ADusername`nPassword: $ADpassword`n`n"
    }

    # checks for IPM username and adds to liquidfiles
    if ($IPMusername) {
    $messageText += "IPM Account`nUsername: $IPMusername`nPassword: $IPMpassword`n`n"
    }

    # checks for EMR username and adds to liquidfiles
    if ($EMRusername) {
    $messageText += "EMR Account`nUsername: $EMRusername`nPassword: $EMRpassword`n`n"
    }

    # checks for EMR username and adds to liquidfiles
    if ($BOSSNETusername) {
        $messageText += "BOSSNET Account`nUsername: $BOSSNETusername`nPassword: $BOSSNETpassword`n`n"
    }

    # this is what is contained in the liquidfile
    $body = @{
        "message" = @{
            "recipients" = @($email)
            "subject"    = "Account Creation - $($ServiceRequest)"
            "message"    = $messageText
            "send_email" = $true
            "private_message" = $true
        }
    } | ConvertTo-Json -Depth 3

    try {
    $response = Invoke-WebRequest `
        -Uri "https://liquidfiles.gha.net.au/message" `
        -Method POST `
        -Headers @{
            Authorization = "Basic $encodedCredentials"
            Accept        = "application/json"
            "Content-Type" = "application/json"
        } `
        -Body $body
    } catch {
        $response = New-Object PSObject
        $response | Add-Member NoteProperty StatusCode 0
    }
    # send the correct status to terminal without long error message
    if ($response.StatusCode -eq 200) {
        Write-Host "Liquidfiles sent successfully"
    } elseif ($response.StatusCode -eq 401) {
        Write-Host "Unauthorized, API key or User authentication failed"
    } elseif ($response.StatusCode -eq 422) {
        Write-Host "Something went wrong and the request could not be completed (Email and/or password incorrect, invalid file, invalid message)"
    } elseif ($response.StatusCode -eq 500) {
        Write-Host "Something went very wrong. This could happen if the system is expecting a value between 0-100 and you send a 1Mb picture of a cat, and similar situations."
    } else {
        Write-Host "Unknow Error"
    }
}
$filepath = Join-Path -Path (Get-Location) -ChildPath "users.csv"
Remove-Item -Path $filepath -Force
