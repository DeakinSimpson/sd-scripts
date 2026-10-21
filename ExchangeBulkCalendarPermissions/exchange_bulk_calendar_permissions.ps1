param (
    [string]$Filepath = "nullptr",
    [string]$Action = "nullptr",
    [string]$Access = "nullptr",
    [string]$ExchangeEmail = "nullptr"
)
 
# ----- GLOBAL DATA -----
 
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
 
function CheckExcelModule
{
    [CmdletBinding()]
    param ()
 
    Write-Host "Starting Excel Module Check..." -ForegroundColor Magenta
 
    # try to import excel sheet module
    if (-not (Get-Module -ListAvailable -Name ImportExcel))
    {
        try
        {
            Write-Host "User Does Not Have ImportExcel Installed, Installing Now..."
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
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
}
 
function CheckFilePathExists
{
    Write-Host "Checking Filepath $($Filepath)..." -ForegroundColor Magenta
 
    if ($Filepath -eq "nullptr")
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
   
    Write-Host "Opening Excel Sheet..." -ForegroundColor Magenta
 
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
 
    Write-Host "Getting Users and Calendars from Excel Sheet" -ForegroundColor Magenta
 
    # get the users and calendars from the excel sheet
    $Users = $excel_sheet.UserEmail
    $Calendars = $excel_sheet.CalendarEmail
 
    # check that the excel sheet has the correct headers
    if (($null -eq $Users) -or ($null -eq $Calendars))
    {
        Write-Warning "Invalid Excel Sheet Format Given, Sheet Must Contain at Least One User and Email, Column Headings MUST be: (UserEmail, CalendarEmail)"
        throw "Invalid Excel Sheet Format Process Halted"
    }
 
    [PSCustomObject]@{
        Users = $Users
        Calendars = $Calendars
    }
}
 
function GetActionFromUser
{
    Write-Host "Please Select one of the following option (input integer):"
 
    [int]$choice = -1
 
    for ($i = 0; $i -le $ValidActionTypes.Count - 1; $i++)
    {
        Write-Host "    $($i) $($ValidActionTypes[$i])"
    }
 
    try
    {
        $choice = Read-Host -Prompt ">"
 
        if (($choice -lt 0) -or ($choice -ge $ValidActionTypes.Count))
        {  
            Write-Warning "Value Provided not on list"
            throw "Number Provided Out of Range Process Halted"
        }
 
        return $ValidActionTypes[$choice]
    }
    catch
    {
        Write-Warning "Invalid Input"
        throw "Invalid Input Process Halted"
    }
 
    return [string]""
}
 
function VerifyAndReturnAction
{
    [CmdletBinding()]
    param ([string]$_Action)
 
    Write-Host "Starting to Verify And Return Action" -ForegroundColor Magenta
 
    if ($_Action -eq "nullptr")
    {
        $_Action = GetActionFromUser
    }
 
    # check if the Action type provided is valid
    if (-not ($_Action -in $ValidActionTypes))
    {
        Write-Warning "Invalid Action Provided"
        throw "Invalid Action Provided Process Halted"
    }
 
    return $_Action
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
 
        return $ValidAccessTypes[$choice]
    }
    catch
    {
        Write-Warning "Invalid Input"
        throw "Invalid Input Process Halted"
    }
 
    return [string]""
}
 
function VerifyAndReturnAccess
{
    param([string]$_Access, [string]$_Action)
 
    Write-Host "Starting to Verify And Return Access" -ForegroundColor Magenta
 
    # Check that Access was proveded for AddAccess type
    if (($_Action -eq "AddAccess") -and ($_Access -eq "nullptr"))
    {
        $_Access = GetAccessFromUser
    }
 
    # if adding access, check that the access type provided is valid
    if (($_Action -eq "AddAccess") -and (-not ($_Access -in $ValidAccessTypes)))
    {
        Write-Warning "Invalid Access Type Provided"
        throw "Invalid Access Type Provieded Process Halted"
    }
 
    return $_Access
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
 
    $exchangeEmail = Read-Host -Prompt "<"
 
    return $exchangeEmail;
}
 
function VerifyAndGetExchangeEmail
{
    param([string]$_ExchangeEmail)
 
    # check that exchange email was provided
    if ($ExchangeEmail -eq "nullptr")
    {
        $_ExchangeEmail = GetExchangeEmail
    }
 
    return $_ExchangeEmail
}
 
function EstablishExchangeConnection
{
    param([string]$_exchangeEmail)
 
    CheckAndInstallExchangeModule
 
    # import the exchange module
    Import-Module ExchangeOnlineManagement
 
    $_exchangeEmail = VerifyAndGetExchangeEmail $_exchangeEmail
 
    # connect via exchange email
    try
    {
        Write-Host "Connecting to Echange Online as $($_exchangeEmail)..." -ForegroundColor Green
        Connect-ExchangeOnline -UserPrincipalName $_exchangeEmail -ShowBanner:$false
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
 
function RemoveCalendarAccess
{
    param($_users, $_calendars)
 
    foreach ($cal in $_calendars)
    {
        if ($cal -eq $null) { continue }
 
        Write-Host "`nRemoving Access to $($cal) Calendar:"
 
        foreach ($user in $_users)
        {
            if ($user -eq $null) { continue }
 
            try {
                Remove-MailboxFolderPermission -Identity "$($cal):\Calendar" -User "$($user)" -Confirm:$false -ErrorAction Stop
                Write-Host "$($user) Successfully had Permissions to $($cal) Calendar Removed" -ForegroundColor Green
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
 
function RemoveDeletedItemsAccess
{
    param($_users, $_calendars)
 
    foreach ($cal in $_calendars)
    {
        if ($cal -eq $null) { continue }
 
        Write-Host "`nRemoving Access to $($cal) Deleted Items:"
 
        foreach ($user in $_users)
        {
            if ($user -eq $null) { continue }
 
            try {
                Remove-MailboxFolderPermission -Identity "$($cal):\Deleted items" -User "$($user)" -Confirm:$false -ErrorAction Stop
                Write-Host "$($user) Successfully had Permissions to $($cal) Deleted Items Removed" -ForegroundColor Green
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
 
    RemoveCalendarAccess $_users $_calendars
    RemoveDeletedItemsAccess $_users $_calendars
}
 
function AddPermission {
    param(
        [string]$_identity,
        [string]$_user,
        [string]$_accessRights
    )
    try {
        Add-MailboxFolderPermission -Identity $_identity -User $_user -AccessRights $_accessRights -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -like "*existing permission entry*" -or $_.ToString() -like "*existing permission entry*") {
            Write-Host "   Already Has Access to: $_identity" -ForegroundColor Yellow
        }
        else {
            Write-Host "   Uncaught Error on $($_identity): $_" -ForegroundColor Red
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
                AddPermission "$($cal):\Deleted Items" "$($user)" Editor
            } elseif ($_access -eq "EditorAndDeleted")
            {
                # add editor permisions
                AddPermission "$($cal):\Calendar" "$($user)" Editor
 
                # add deleted items folder permisions
                AddPermission "$($cal):\Deleted Items" "$($user)" Editor
            } else
            {
                AddPermission "$($cal):\Calendar" "$($user)" Editor
            }
        }
    }
}
 
# ---- CODE -----
 
# get the excel sheet rows
$ExcelSheetResult = GetUsersAndCalendars $Filepath
$Users = $ExcelSheetResult.Users
$Calendars = $ExcelSheetResult.Calendars
 
# get what needs to be performed to the mailboxes
$Action = VerifyAndReturnAction $Action
$Access = VerifyAndReturnAccess $Access $Action
 
# connect to exchange
EstablishExchangeConnection $ExchangeEmail
 
# check each user email exists
Write-Host "`nChecking Each User Has an Exchange Online Mailbox..."
CheckMailboxesExist $Users
 
# check each calendar email exists
Write-Host "`nChecking Each Calender Has an Exchange Online Mailbox"
CheckMailboxesExist $Calendars
 
# run excecution query for all users and all calendars
    # log all action results
 
if ($Action -eq "CheckAccess")
{
    CheckAccess $Users
}
 
if ($Action -eq "RemoveCalendarAccess")
{
    RemoveCalendarAccess $Users $Calendars
}
 
if ($Action -eq "RemoveDeletedItemsAccess")
{
    RemoveDeletedItemsAccess $Users $Calendars
}
 
if ($Action -eq "RemoveAllAccess")
{
    RemoveAllAccess $Users $Calendars
}
 
if ($Action -eq "AddAccess")
{
    AddAccess $Users $Calendars $Access
}
 
# TODO:
    # when voiding a parameter, prompt the user to enter it and provide a list
    # try and implement display names rather then emails
    # try and implement display name searching for calendars rather then email
    # allow for the option to only select one email (no csv)
