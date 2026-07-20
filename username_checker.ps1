# imports a csv from a path and checks for errors
function importCSV($csv_path) {
    # check if path exists
    if (-not (Test-Path -Path $csv_path -PathType Leaf -ErrorAction SilentlyContinue)) {
        Write-Host "`n`$($csv_path) does not exist"
        return $null
    }

    # import csv
    $csv = Import-Csv -Path $csv_path

    # check if there is any data
    if ($null -eq $csv) {
        Write-Host "csv contains no data"
        return $null
    }

    return $csv
}

function convertFirstAndLastToUsername($first_and_last_csv) {
    $usernames = @()

    # for each first and last name pair
    foreach ($row in $first_and_last_csv) {
        # combine into "first.last"
        $username = $row.first_name + "." + $row.last_name

        # append to usernames list
        $usernames += $username
    }

    # return the list of first.last usernamesS
    return $usernames
}

# checks a list of usernames
function checkUsernames($usernames) {
    # initialise the empty lists
    $users_with_ad = @()
    $users_without_ad = @()

    # loop through all usernames in the csv
    foreach ($user in $usernames) {
        # add user to the correct list
        if (Get-ADUser -Filter "SamAccountName -eq '$user'" -ErrorAction SilentlyContinue) {
            $users_with_ad += $user
        } else {
            $users_without_ad += $user
        }
    }

    # return an object of containing the first and last name
    return [PSCustomObject]@{
        users_with_ad = $users_with_ad
        users_without_ad = $users_without_ad
    }
}

# this is the list of OU names that disabled accounts sit in, make sure to make them the all lowercase version
$disabled_ous = @(
        "disabled",
        "disabled users",
        "terminated"
    )

# checks if each user is disabled
function checkUserDisabled($usernames) {
    $user_disable_status = @()

    # loop through all usernames
    foreach ($user in $usernames) {
        $is_user_disabled = $false

        # get the OU of the user
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$user'" -Properties DistinguishedName

        if ($null -eq $adUser) {
            Write-Warning "Couldn't find user '$user'"
            $user_disable_status += $false
            continue
        }


        # get the string of the OU
        $ou_string = ($adUser.DistinguishedName -replace '^CN=.*?,') -replace ',DC=.*$'

        # split ou sting at ',' giveing an array of split stings
        $ou_split = $ou_string -split ","

        # loop through each character in the split string
        foreach ($part in $ou_split) {
            # remove the 'OU=' prefix
            $part_removed_prefix = $part -replace "OU="

            # convert to lowercase
            $part_to_lower = $part_removed_prefix.ToLower()

            # check each disabled term
            foreach ($disabled_ou in $disabled_ous) {
                # if the account is in one of the disabled ous, mark as disabled
                if ($part_to_lower -eq $disabled_ou) {
                    $is_user_disabled = $true
                }
            }
        }

        $user_disable_status += $is_user_disabled
    }

    return $user_disable_status
}

# print functions
# -------------------------------------------------------------

function printTitleCard {
    Write-Host "------------------------------------------------------------"
    Write-Host "|                   CSV username checker                   |"
    Write-Host "------------------------------------------------------------"
}

function printPathMenu {
    Write-Host "`n`Please Enter the Path of the .csv, path\to\file.csv:"
}

function printCSVSelectMenu {
    Write-Host "`n`Please Select From the Following Options, then press ENTER"
    Write-Host "    1) CSV with Usernames"
    Write-Host "    2) CSV with First & Last Name"

}

# print the users
function printAdUsers($list, $disabled_users) {
    for ($i = 0; $i -lt $list.Count; $i++) {
        Write-Host "$($list[$i]) - $($disabled_users[$i])"
    }
}

function printNoAdUsers($list) {
    foreach ($item in $list) {
        Write-Host "$($item)"
    }
}

# prints the usernames that do and dont have AD accounts
function printFullResult($result, $disabled_users) {
    # print each user with AD account
    Write-Host "`n`Users With AD:"
    [void](printAdUsers $result.users_with_ad $disabled_users)

    # print users without AD
    Write-Host "`n`Users Without AD:"
    [void](printNoAdUsers $result.users_without_ad)
}

# switch case functions
# -------------------------------------------------------------

# gets the csv path
function getCSVPath {
    [void](printPathMenu)

    return Read-Host -Prompt " > "
}

function getCSVType {
    [void](printCSVSelectMenu)

    $user_input = Read-Host -Prompt " > "

    switch ($user_input) {
        1 { return 2 }
        2 { return 3 }
        default {
            Write-Host "Invalid Input, Please Try Again!"
            return 1
        }
    }
}

function runUsernameCheck($csv_path) {
    $csv = importCSV($csv_path)

    # only run if importCSV is successfull
    if ($null -ne $csv) {
        $usernames = $csv.usernames

        $result = checkUsernames $usernames

        $disabled_status = checkUserDisabled $result.users_with_ad

        printFullResult $result $disabled_status
    }
}

function runNameCheck($csv_path) {
    $names = importCSV($csv_path)

    # only run if importCSV is successfull
    if ($null -ne $names) {
        $usernames = convertFirstAndLastToUsername $names

        $result = checkUsernames $usernames

        $disabled_status = checkUserDisabled $result.users_with_ad

        printFullResult $result $disabled_status
    } 
}


# main function
# -------------------------------------------------------------

function main {
    Clear-Host
    [void](printTitleCard)

    # this is used to change what part of the switch we are on
    $case = 0

    # this is the location of the csv_file
    $csv_path

    while ($true){
        switch ($case) {
            0 {
                # gets the csv path
                $csv_path = getCSVPath
                $case = 1
            }
            1 {
                # gets the method used
                $case = getCSVType
            }
            2 {
                # runs for usernames
                runUsernameCheck($csv_path)
                $case = 4
            }
            3 {
                # runs for first and last name
                runNameCheck $csv_path
                $case = 4
            }
            4 {
                # exit case
                Write-Host "`n`Exiting Program"
                return
            }
            default {
                Write-Host "Invalid Menu, Sending to Main"
                $case = 0
            }
        }
    }
}

main
