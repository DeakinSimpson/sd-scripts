function checkUsernames([string]$csv_path) {
    # get the usernames from the csv
    $usernames = Import-Csv -Path $csv_path

    # initialise the empty lists
    $users_with_ad = @()
    $users_without_ad = @()

    # loop through all usernames in the csv
    foreach ($row in $usernames) {
        $user = $row.usernames

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
        Write-Host " > $($item)"
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
                # prompts user to enter csv path
                [void](printPathMenu)

                $csv_path = Read-Host -Prompt " > "

                $case = 1
            }
            1 {
                # prompts user to enter csv type
                [void](printCSVSelectMenu)
                $user_input = Read-Host -Prompt " > "

                if      ($user_input -eq 1) {$case = 2}
                elseif  ($user_input -eq 2) {$case = 3}
                else                        {Write-Host "Invalid Input, Please Try Again!"}
                
            }
            2 {
                # this is the case for using usernames
                $result = checkUsernames($csv_path)

                printFullResult($result)

                $case = 4

            }
            3 {
                # this is the case if using first and last name
                Write-Host "`n`This is not implemented yet, sending to main..."
                $case = 1
            }
            4 {
                # exit case
                Write-Host "Exiting Program"
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
