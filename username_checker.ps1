function getCSVUsernames([string]$csv_path) {
    # get the usernames from the csv
    $usernames = Import-Csv -Path $csv_path

    # returns a list of usernames
    return $usernames.usernames
}

function getFirstAndLastNames($csv_path) {
    $csv = Import-Csv -Path $csv_path

    # returns the raw csv with first and last names
    return $csv
}

function convertFirstAndLastToUsername($first_and_last_csv) {
    $usernames = @()

    foreach ($row in $first_and_last_csv) {
        $username = $row.first_name + "." + $row.last_name
        $usernames += $username
    }

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

    return [PSCustomObject]@{
        users_with_ad = $users_with_ad
        users_without_ad = $users_without_ad
    }
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

function printList($list) {
    foreach ($item in $list) {
        Write-Host "$($item)"
    }
}

function printFullResult($result) {
    # print each user with AD account
    Write-Host "`n`Users With AD:"
    [void](printList($result.users_with_ad))

    # print users without AD
    Write-Host "`n`Users Without AD:"
    [void](printList($result.users_without_ad))
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
    $usernames = getCSVUsernames($csv_path)

    $result = checkUsernames($usernames)

    printFullResult($result)
}

function runNameCheck($csv_path) {
    $names = getFirstAndLastNames($csv_path)

    $usernames = convertFirstAndLastToUsername($names)

    $result = checkUsernames($usernames)

    printFullResult($result)
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
                $csv_path = getCSVPath
                $case = 1
            }
            1 {
                $case = getCSVType
            }
            2 {
                runUsernameCheck $csv_path
                $case = 4
            }
            3 {
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
