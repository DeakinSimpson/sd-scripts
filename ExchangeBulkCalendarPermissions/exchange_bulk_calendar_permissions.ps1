
param (
    [string]$Filepath = "NotProvided",
    [string]$Action = "NotProvided",
    [string]$Access = "NotProvided",
    [string]$ExchangeEmail = "NotProvided",
    [string]$Mode = "default"
)
 
# ----- GLOBAL DATA -----
$WindowWidth = [Console]::WindowWidth
$Users
$Calendars
 
function DebugOut
{
    param([string] $output)
 
    if ($Mode -eq "Verbose")
    {
        Write-Host $output -ForegroundColor Magenta
    }
}
 
function WriteJsonLog
{
    param(
        [string]$Message,
        [hashtable]$Details = @{},
        [string]$Path = "C:\logs\exchange_bulk_logs.json"
    )
 
    $LogFolder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -Path $LogFolder))
    {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }
 
    $LogEntry = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Message = $Message
        Metadata  = $Details
        ScriptUser = $env:USERNAME
    }
 
    $LogEntry | ConvertTo-Json -Depth 10 -Compress | Add-Content -Path $Path -Encoding utf8
}
 
# array of all valid access types
$ValidActionTypes = @(
    "CheckAccess"
    "RemoveCalendarAccess"
    "RemoveDeletedItemsAccess"
    "RemoveAllAccess"
    "AddAccess"
)
 
# Consult https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-mailboxfolderpermission?…
$ValidAccessTypes = @(
    "None"              # --- Individual Permisions ---
    "CreateItems"
    "CreateSubfolders"
    "DeleteAllItems"
    "DeleteOwnedItems"
    "EditAllItems"
    "EditOwnedItems"
    "FolderContact"
    "FolderOwner"
    "FolderVisible"
    "ReadItems"
    "Author"            # --- Roles ---
    "Contributor"
    "Editor"
    "NonEditingAuthor"
    "Owner"
    "PublishingAuthor"
    "PublishingEditor"
    "Reviewer"
    "AvailabilityOnly"  # --- Calendar Only Permissions ---
    "LimitedDetails"
    "DeletedItems"      # --- Custom Functions ---
    "EditorAndDeleted"
)
 
# ----- FUNCTIONS -----
 
function GetLineForHeading
{
    param([string]$Name, [string]$Value)
 
    DebugOut "Starting to Get Line For Heading with Name: $($Name), Value: $($Value)"
 
    $Text = "| $($Name): " + $Value
    $TextLine = $Text + (" " * ($WindowWidth - $Text.Length - 1)) + "|"
 
    Write-Host $TextLine
}
 
function WriteLine
{
    DebugOut "Starting to Write Line at Terminal Width: $($WindowWidth)"
 
    $Line = "=" * $WindowWidth
    Write-Host "$Line"
}
 
function WriteTitleRow
{
    $WindowWidth = [Console]::WindowWidth
    $Title = "Exchange Bulk Calendar Script"
 
    $TotalInnerSpace = $WindowWidth - 2
    $LeftPadding = [math]::Max(0, [math]::Floor(($TotalInnerSpace - $Title.Length) / 2))
    $RightPadding = [math]::Max(0, $TotalInnerSpace - $Title.Length - $LeftPadding)
 
    $TitleRow = "|" + (" " * $LeftPadding) + $Title + (" " * $RightPadding) + "|"
 
    Write-Host $TitleRow
}
 
function WriteHeading
{
    DebugOut "Writing title heading..."
 
    if ($Mode -ne "Verbose") { Clear-Host }
 
    WriteLine
    WriteTitleRow
    WriteLine
    GetLineForHeading "Filepath" $Filepath
    GetLineForHeading "Action" $Action
    GetLineForHeading "Access" $Access
    GetLineForHeading "Exchange Email" $ExchangeEmail
    GetLineForHeading "Mode" $Mode
    WriteLine
}
 
function CheckExcelModule
{
    [CmdletBinding()]
    param ()
 
    DebugOut "Starting Excel Module Check..."
 
    # try to import excel sheet module
    if (-not (Get-Module -ListAvailable -Name ImportExcel))
    {
        try
        {
            Write-Host "User Does Not Have ImportExcel Installed, Installing Now..."
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
 
            DebugOut "Module Successfully Installed."
        }
        catch
        {
            Write-Warning "Cant Import Module Required for Excel Sheet Importing"
            throw "Import Module Failed Process Halted"
        }
    }
}
 
function GetFilepathFromUser
{
    Write-Host "Filepath was not provided, please provide a VALID filepath for the excel sheet to be processed:"
 
    $script:Filepath = Read-Host -Prompt ">"
 
    DebugOut "Filepath retrieved from user..."
 
    WriteHeading
}
 
function CheckFilePathExists
{
    DebugOut "Checking Filepath $($Filepath)..."
 
    if ($Filepath -eq "NotProvided")
    {
        GetFilepathFromUser
    }
 
    # check if the filepath provided exists
    if (-not (Test-Path -Path $Filepath))
    {
        Write-Warning "Failed to find path: $($Filepath)"
        throw "File Path Not Found Process Halted"
    }
}
 
function GetExcelSheet
{
    [CmdletBinding()]
    param ($_Filepath)
   
    DebugOut "Opening Excel Sheet..."
 
    # try to open the file path
    try
    {
        $excel_sheet = Import-Excel -Path $Filepath
    }
    catch
    {
        Write-Warning "Failed to open excel sheet at file path: $($Filepath)"
        throw "Excel Sheet Failed to Open Process Halted"
    }
 
    return $excel_sheet
}
 
function GetUsersAndCalendars
{
    # check modules and pull excel sheet
    CheckExcelModule
    CheckFilePathExists $Filepath
    $excel_sheet = GetExcelSheet $Filepath
 
    DebugOut "Getting Users and Calendars from Excel Sheet"
 
    # get the users and calendars from the excel sheet
    $Users = $excel_sheet.UserEmail
    $Calendars = $excel_sheet.CalendarEmail
 
    # check that the excel sheet has the correct headers
    if (($null -eq $Users) -or ($null -eq $Calendars))
    {
        Write-Warning "Invalid Excel Sheet Format Given, Sheet Must Contain at Least One User and Email, Column Headings MUST be: (UserEmail, CalendarEmail)"
        throw "Invalid Excel Sheet Format Process Halted"
    }
 
    DebugOut "Returning Users and Calenders as PSCustomObject"
 
    [PSCustomObject]@{
        Users = $Users
        Calendars = $Calendars
    }
}
 
function GetActionFromUser
{
    DebugOut "Action was not provided, requesting action from user"
 
    Write-Host "`nPlease Select one of the following option (input integer):"
 
    [int]$choice = -1
 
    DebugOut "Starting to display each Action available"
    for ($i = 0; $i -le $ValidActionTypes.Count - 1; $i++)
    {
        Write-Host "    $($i) $($ValidActionTypes[$i])"
    }
 
    try
    {
        $choice = Read-Host -Prompt ">"
       
        DebugOut "Input was a number, checking to see if it is in range"
        if (($choice -lt 0) -or ($choice -ge $ValidActionTypes.Count))
        {  
            Write-Warning "Value Provided not on list"
            throw "Number Provided Out of Range Process Halted"
        }
 
        DebugOut "Updating $script:Action"
        $script:Action = $ValidActionTypes[$choice]
 
        WriteHeading
    }
    catch
    {
        Write-Warning "Invalid Input"
        throw "Invalid Input Process Halted"
    }
}
 
function VerifyAction
{
    DebugOut "Starting to Verify And Return Action"
 
    if ($Action -eq "NotProvided") { GetActionFromUser }
 
    # check if the Action type provided is valid
    if (-not ($script:Action -in $ValidActionTypes))
    {
        Write-Warning "Invalid Action Provided"
        throw "Invalid Action Provided Process Halted"
    }
}
 
function GetAccessFromUser
{
    Write-Host "Please Select one of the following option (input integer):"
 
    [int]$choice = -1
 
    for ($i = 0; $i -le $ValidAccessTypes.Count - 1; $i++)
    {
        Write-Host "    $($i) $($ValidAccessTypes[$i])"
    }
 
    try
    {
        $choice = Read-Host -Prompt ">"
 
        if (($choice -lt 0) -or ($choice -ge $ValidAccessTypes.Count))
        {  
            Write-Warning "Value Provided not on list"
            throw "Number Provided Out of Range Process Halted"
        }
 
        $script:Access = $ValidAccessTypes[$choice]
 
        WriteHeading
    }
    catch
    {
        Write-Warning "Invalid Input"
        throw "Invalid Input Process Halted"
    }
}
 
function VerifyAndReturnAccess
{
    DebugOut "Starting to Verify And Return Access"
 
    # Check that Access was proveded for AddAccess type
    if (($Action -eq "AddAccess") -and ($Access -eq "NotProvided"))
    {
        GetAccessFromUser
    }
 
    # if adding access, check that the access type provided is valid
    if (($Action -eq "AddAccess") -and (-not ($Access -in $ValidAccessTypes)))
    {
        Write-Warning "Invalid Access Type Provided"
        throw "Invalid Access Type Provieded Process Halted"
    }
}
 
function CheckAndInstallExchangeModule
{
    # check that modules are installed, if not install
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement))
    {
        Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Cyan
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    }
}
 
function GetExchangeEmail
{
    Write-Host "Exchange Email not provided, please provide an email to connect to the exchange server: "
 
    $script:ExchangeEmail = Read-Host -Prompt ">"
 
    WriteHeading
}
 
function VerifyAndGetExchangeEmail
{
    # check that exchange email was provided
    if ($ExchangeEmail -eq "NotProvided")
    {
        GetExchangeEmail
    }
}
 
function EstablishExchangeConnection
{
    CheckAndInstallExchangeModule
 
    # import the exchange module
    Import-Module ExchangeOnlineManagement
 
    VerifyAndGetExchangeEmail
 
    # connect via exchange email
    try
    {
        Write-Host "Connecting to Echange Online as $($ExchangeEmail)..." -ForegroundColor Green
        Connect-ExchangeOnline -UserPrincipalName $ExchangeEmail -ShowBanner:$false
        WriteJsonLog -Message "Connect to Exchange" -Details @{User = $ExchangeEmail}
        Write-Host "Successfully Connected to Exchange Online." -ForegroundColor Green
    }
    catch
    {
        Write-Warning "Failed to Connect to Exchange Online. Please check that the email provided is valid."
        throw "Failed to Connect to Exchange Online Process Halted"
    }
}
 
function CheckMailboxesExist
{
    param($_mailboxes)
 
    foreach ($mailbox in $_mailboxes)
    {
        if ($mailbox -eq $null) { continue }
 
        WriteJsonLog -Message "Check Mailbox" -Details @{
            MailboxBeingChecked = "$mailbox"
            CommandRun = "Get-Recipient -Identity $mailbox -ErrorAction SilentlyContinue"
        }
        $recipient = Get-Recipient -Identity $mailbox -ErrorAction SilentlyContinue
 
        if ($recipient -and $recipient.RecipientTypeDetails -match "Mailbox")
        {
            Write-Host "$($mailbox) Mailbox Found and Verified :)" -ForegroundColor Green
        }
        else
        {
            # It found a contact, group, or nothing at all
            Write-Host "$($mailbox) is NOT a valid mailbox (Type: $($recipient.RecipientTypeDetails))!!!" -ForegroundColor Red
            throw "Invalid Mailbox Process Halted"
        }
    }
}
 
function CheckAccess
{
    param($_users)
 
    foreach ($user in $_users)
    {
        if ($user -eq $null) { continue }
 
        Write-Host "`nAccess to $($user) Calendar:" -ForegroundColor Green
 
        $calendarPermissions = $null
 
        WriteJsonLog -Message "Check Access" -Details @{
            CommandRun = 'Get-MailboxFolderPermission -Identity "$($user):\Calendar" -ErrorAction SilentlyContinue'
            CommandDetails = @{
                UsersAccessWhoWasChecked = $user
                FolderChecked = "Calendar"
            }
        }
        $calendarPermissions = Get-MailboxFolderPermission -Identity "$($user):\Calendar" -ErrorAction SilentlyContinue
 
        if ($calendarPermissions)
        {
            $calendarPermissions | Format-Table FolderName, User, AccessRights
        }
        else
        {
            Write-Host "No Calendar Permissions Found!" -ForegroundColor Red
        }
 
        Write-Host "`nAccess to $($user) Deleted Items:" -ForegroundColor Green
 
        $deletedItemsPermissions = $null
 
        WriteJsonLog -Message "Check Access" -Details @{
            CommandRun = 'Get-MailboxFolderPermission -Identity "$($user):\Deleted Items" -ErrorAction SilentlyContinue'
            CommandDetails = @{
                UsersAccessWhoWasChecked = $user
                FolderChecked = "Deleted Items"
            }
        }
        $deletedItemsPermissions = Get-MailboxFolderPermission -Identity "$($user):\Deleted Items" -ErrorAction SilentlyContinue
 
        if ($deletedItemsPermissions)
        {
            $deletedItemsPermissions | Format-Table FolderName, User, AccessRights
        }
        else
        {
            Write-Host "No Deleted Items Permissions Found!" -ForegroundColor Red
        }
 
 
    }
}
 
function RemoveFolderAccess
{
    param($_users, $_calendars, $_folder)
 
    foreach ($cal in $_calendars)
    {
        if ($cal -eq $null) { continue }
 
        Write-Host "`nRemoving Access to $($cal) $($_folder):"
 
        foreach ($user in $_users)
        {
            if ($user -eq $null) { continue }
 
            try {
                WriteJsonLog -Message "Removed Access" -Details @{
                    CommandRun = 'Remove-MailboxFolderPermission -Identity "$($cal):\$($_folder)" -User "$($user)" -Confirm:$false -ErrorAction Stop'
                    CommandDetails = @{
                        UserWhoHadAccessRemoved = $user
                        UserWhoIsTheOwnerOfTheFolder = $cal
                        Folder = $_folder
                    }
                }
                Remove-MailboxFolderPermission -Identity "$($cal):\$($_folder)" -User "$($user)" -Confirm:$false -ErrorAction Stop
                Write-Host "$($user) Successfully had Permissions to $($cal) $($_folder) Removed" -ForegroundColor Green
            }
            catch {
                $errorString = $_.ToString()
 
                if (($errorString -like "*no existing permission*") -or
                    ($errorString -like "*ManagementObjectNotFoundException*"))
                {
                    Write-Host "$($user) Did not have permissions to $($cal)" -ForegroundColor Yellow
                }
                else
                {
                    Write-Host "A different error occurred: $_" -ForegroundColor Red
                    throw "Failed To Remove Permissions Process Halted"
                }
            }
        }
    }
}
 
function RemoveAllAccess
{
    param($_users, $_calendars)
 
    RemoveFolderAccess $_users $_calendars "Calendar"
    RemoveFolderAccess $_users $_calendars "Deleted Items"
}
 
function AddPermission {
    param(
        [string]$_folder_user,
        [string]$_folder,
        [string]$_user,
        [string]$_accessRights
    )
    try {
        WriteJsonLog -Message "Add Access" -Details @{
            CommandRun = 'Add-MailboxFolderPermission -Identity "$($_folder_user):\$($_folder)" -User $_user -AccessRights $_accessRights -ErrorAction Stop'
            CommandDetails = @{
                UserWhoGainedAccess = $_user
                OwnerOfFolderGivenAccessTo = $_folder_user
                Folder = $_folder
                AccessRights = $_accessRights
            }
        }
        Add-MailboxFolderPermission -Identity "$($_folder_user):\$($_folder)" -User $_user -AccessRights $_accessRights -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -like "*existing permission entry*" -or $_.ToString() -like "*existing permission entry*") {
            Write-Host "   Already Has Access to: $_folder_user" -ForegroundColor Yellow
        }
        else {
            Write-Host "   Uncaught Error on $($_folder_user): $_" -ForegroundColor Red
            throw
        }
    }
}
 
function AddAccess
{
    param($_users, $_calendars, $_access)
 
    foreach ($cal in $_calendars)
    {
        if ($null -eq $cal) { continue }
 
        Write-Host "Adding Access to $($cal): " -ForegroundColor Blue
 
        foreach ($user in $_users)
        {
            if ($null -eq $user) { continue }
 
            Write-Host "$($user) is gaining $($_access) permissions to $($cal)" -ForegroundColor Blue
           
            if ($_access -eq "DeletedItems")
            {
                AddPermission $cal "Deleted Items" $user Editor
            } elseif ($_access -eq "EditorAndDeleted")
            {
                # add editor permisions
                AddPermission $cal "Calendar" $user Editor
 
                # add deleted items folder permisions
                AddPermission $cal "Deleted Items" $user Editor
            } else
            {
                AddPermission $cal "Calendar" $user Editor
            }
        }
    }
}
 
# ---- CODE -----
WriteHeading
 
# connect to exchange
EstablishExchangeConnection
 
# get the excel sheet rows
$ExcelSheetResult = GetUsersAndCalendars $Filepath
$Users = $ExcelSheetResult.Users
$Calendars = $ExcelSheetResult.Calendars
 
# get what needs to be performed to the mailboxes
VerifyAction
VerifyAndReturnAccess
 
 
 
# check each user email exists
Write-Host "`nChecking Each User Has an Exchange Online Mailbox..."
CheckMailboxesExist $Users
 
# check each calendar email exists
Write-Host "`nChecking Each Calender Has an Exchange Online Mailbox"
CheckMailboxesExist $Calendars
 
# run excecution query for all users and all calendars
    # log all action results
 
switch ($Action)
{
    "CheckAccess"               { CheckAccess $Users }
    "RemoveCalendarAccess"      { RemoveCalendarAccess $Users $Calendars }
    "RemoveDeletedItemsAccess"  { RemoveDeletedItemsAccess $Users $Calendars }
    "RemoveAllAccess"           { RemoveAllAccess $Users $Calendars }
    "AddAccess"                 { AddAccess $Users $Calendars $Access }
}
 
# TODO:
    # try and implement display names rather then emails
    # try and implement display name searching for calendars rather then email
    # allow for the option to only select one email (no csv)
