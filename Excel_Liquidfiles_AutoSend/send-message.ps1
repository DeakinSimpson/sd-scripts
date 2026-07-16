# variables for api key
# grab the api key from liquidfiles
$apikey = ""
# credential encoding
$pair = "$($apikey):$()"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
# import the csv
$csvPath = Join-Path -Path (Get-Location) -ChildPath "users.csv"
$users = Import-Csv -Path $csvPath
# used to identify what line successfully sent and what line had an error
$lineCount = 2
$failedLines = @()
foreach ($user in $users) {
    $messageText        = ""
    $email              = $user.'requesters email'
    $ADusername         = $user.'AD username'
    $ADpassword         = $user.'AD password'
    $IPMusername        = $user.'IPM username'
    $IPMpassword        = $user.'IPM password'
    $EMRusername        = $user.'EMR username'
    $EMRpassword        = $user.'EMR password'
    $BOSSNETusername    = $user.'BOSSNET username'
    $BOSSNETpassword    = $user.'BOSSNET password'
    $ServiceRequest     = $user.'SR'
    $firstname          = $user.'firstname'
    $surname            = $user.'surname'
    # checks if email is epmpty, skips if theres none and gives output
    if (-not $email -or $email.Trim() -eq "") {
        Write-Host "Line $($lineCount): Skipping row with no email"
        $failedLines += $lineCount
        $lineCount++
        continue
    }
    # checks each account and adds them to the message if they exist
    if ($firstname -and $surname) {$messageText += "$($firstname) $($surname)`n`n"}
    if ($ADusername) {$messageText      += "Computer Account`nUsername: $ADusername`nPassword: $ADpassword`n`n"}
    if ($IPMusername) {$messageText     += "IPM Account`nUsername: $IPMusername`nPassword: $IPMpassword`n`n"}
    if ($EMRusername) {$messageText     += "EMR Account`nUsername: $EMRusername`nPassword: $EMRpassword`n`n"}
    if ($BOSSNETusername) {$messageText += "BOSSNET Account`nUsername: $BOSSNETusername`nPassword: $BOSSNETpassword`n`n"}
    # this is what is contained in the liquidfile
    $body = @{
        "message" = @{
            "recipients"        = @($email)
            "subject"           = "Account Creation - $($firstname) $($surname) - $($ServiceRequest)"
            "message"           = $messageText
            "send_email"        = $true
            "private_message"   = $true
        }
    } | ConvertTo-Json -Depth 3
    # this tries to invoke the web request
    try {
    $response = Invoke-WebRequest `
        -Uri "https://liquidfiles.gha.net.au/message" `
        -Method POST `
        -Headers @{
            Authorization   = "Basic $encodedCredentials"
            Accept          = "application/json"
            "Content-Type"  = "application/json"
        } `
        -Body $body
    } catch {
        $errorMessage = "Line $($lineCount): $($_.Exception.Message)"
        $failedLines += $lineCount
        $response = New-Object PSObject
        $response | Add-Member NoteProperty StatusCode 0
    }
    # send the correct status to terminal without long error message
    switch ($response.StatusCode) {
        200 { Write-Host "Line $($lineCount): Liquidfiles sent successfully" }
        401 { Write-Host "Unauthorized, API key or User authentication failed" }
        422 { Write-Host "Something went wrong and the request could not be completed (Email and/or password incorrect, invalid file, invalid message)" }
        500 { Write-Host "Something went very wrong. This could happen if the system is expecting a value between 0-100 and you send a 1Mb picture of a cat, and similar situations." }
        default {
            if ($errorMessage) {
                Write-Host "Error: $errorMessage"
            } else {
                Write-Host "Unk nown Error"
            }
        }
    }
    $lineCount++
}
# after the loop display a summary of failed lines (maybe implement showing the service request number or email)
if ($failedLines.Count -gt 0) {
    Write-Host "`nThe following lines failed to send:"
    for ($i = 0; $i -lt $failedLines.Count; $i++) {
        Write-Host "Line $($failedLines[$i])"
    }
} else {
    Write-Host "`nAll lines sent successfully."
}
# delete the csv after it is done so that there is no accidental resending of the same information
$filepath = Join-Path -Path (Get-Location) -ChildPath "users.csv"
Remove-Item -Path $filepath -Force
