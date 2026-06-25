<#
.SYNOPSIS
Removes data from deleted user accounts
.DESCRIPTION
Removes data from deleted user accounts: remove from email groups, mananger, security, hides from Sype and email lists)
.PARAMETER
Prompt - prompt to remove the user
Test - test the script (using whatif) rather than running
.NOTES
Script name: departedUserProcessing.ps1
Version: 1.0
Author: Ryan
Contact: Ryan
DateCreated: 2019
LastUpdate: 2019
#>

[CmdletBinding()]
Param(
[Parameter(
Mandatory=$False)]
[switch]$Prompt=$False,
[Parameter(
Mandatory=$False)]
[switch]$Test=$False
)

<# ### OVERRIDES ### #>
# If you "really" want to test this script in ISE instead of making it a function or, calling the script with parameters
# Uncomment the lines below and set appropriately. !!! COMMENT OUT WHEN DONE !!!
# $Prompt = $False;
# $Test=$True;
<#####################>
$JustTesting = @{
    WhatIf = $Test;
    Verbose = $True;
}
# $Prompt - No need to define and set this variable, it is now a True/False switch
# get user credentials for doing things.
# the could be read from an encrypted file instead of doing this:
if (-not $UserCredential) {
    $UserCredential = get-credential
}

function Move_DisabledUsers {<#
.SYNOPSIS
Removes data from deleted user accounts
.DESCRIPTION
Removes data from deleted user accounts: remove from email groups, mananger, security, hides from Skype and email lists)
.PARAMETER
Prompt - prompt to remove the user
Test - test the script (using whatif) rather than running
.NOTES
Script name: departedUserProcessing.ps1
Version: 1.0
Author: Ryan McArthur
Contact: Ryan.McArthur
DateCreated: 2019
LastUpdate: 2019
#>

[CmdletBinding()]
Param(
[Parameter(
Mandatory=$False)]
[switch]$Prompt=$False,
[Parameter(
Mandatory=$False)]
[switch]$Test=$False
)

<# ### OVERRIDES ### #>
# If you "really" want to test this script in ISE instead of making it a function or, calling the script with parameters
# Uncomment the lines below and set appropriately. !!! COMMENT OUT WHEN DONE !!!
# $Prompt = $False;
# $Test=$True;
<#####################>
$JustTesting = @{
    WhatIf = $Test;
    Verbose = $True;
}
# $Prompt - No need to define and set this variable, it is now a True/False switch
# get user credentials for doing things.
# the could be read from an encrypted file instead of doing this:
if (-not $UserCredential) {
    $UserCredential = get-credential
}
function Move_DisabledUsers {
    #Move disabled users from Users OU to Disabled users OU before we do anything else
    [cmdletbinding()]
    param(
    [Parameter(Mandatory=$True)]
    [string]$fromOU,
    [Parameter(Mandatory=$True)]
    [string]$Server,
    [Parameter(Mandatory=$True)]
    [string]$toOU)
    #$usersToMove = $NULL
    if ($fromOU -eq $null -or $Server -eq $null -or $toOU -eq $null)
    {
        write-host "You need to add the data for function Move_DisabledUsers"
        $returnvalue = $false
    }
    else
    {
        # find disabled users in the OU
        try {
            $usersToMove = get-aduser -searchbase $fromOU -searchscope 1 -filter { enabled -eq $false } -Properties * -Server $Server | sort name
        }
        catch {
            # if it doesn't work, or is empty (no users found) return nothing
            write-host "no users found to move."
            #$returnvalue = $false
        }
        if ($usersToMove) {
            $returnvalue=$null
            foreach ($MoveThisUser in $usersToMove )
            {
                write-host "moving: " $MoveThisUser.Samaccountname
                # check for errors here
                try {
                    move-adobject -identity $MoveThisUser -server $Server -targetpath $toOU @JustTesting
                    # write out a list of the moved users
                    #record the moved users and return to the script
                    #$returnvalue += $usersToMove.name
                }
                catch {
                    write-host "unable to move user:" $error[0]
                    # write an error instead to the list of moved users
                    #$returnvalue += $usersToMove.name +
                }
            }
        }
    }
    # return $returnvalue
}
function process_DisabledUsers {
    # check if the account has already been processed, if not, then process it
    [cmdletbinding()]
    param(
    [Parameter(Mandatory=$True)]
    [string]$DepartedUsersOU,
    [Parameter(Mandatory=$True)]
    [string]$Server)
    write-host $DepartedUsersOU
    $DisabledUserList = get-aduser -searchbase $DepartedUsersOU -filter { enabled -eq $false } -Properties * -Server $Server | sort name
    # change for testing, this just returns one user rather than a huge list.
    # $DisabledUserList = get-aduser -filter { enabled -eq $false -and userPrincipalName -eq <username> } -Properties * -Server $Server
    # use a string to store the results
    $ListProcessedUsers = "Name, info written, manager remove, email hide, skype hide, processed `r`n"
    foreach ( $DisabledUser in $DisabledUserList ) {#1
    write-host " "
    write-host "### Processing " $disableduser.name
    # if Disabled user already processed (ie description starts with "Disabled"
    if ( $DisabledUser.description.contains("Disabled") ) {#2
    write-host "user " $disableduser.name " has already been processed"
}#2
# error reading the notes?
# otherwise:
else { #2
write-host "user " $disableduser.name " has not been processed, processing"
$ListProcessedUsers += $disableduser.name
# reset the variables
$hideSkype = $false
$manager = $false
$email = $false
$hideAddress = $false
# info out into notes make sure it is recorded
write-host "writing information to Notes field from "$Server
$notes = Write_UserInformationToNotes -user $DisabledUser -Server $Server
# read the info out to make sure it is recorded.
if ( Read_UserInformationFromNotes -user $DisabledUser -Server $Server )
{
    $ListProcessedUsers += "," + $notes
}
# prompt to continue if this is a test:
if ($prompt) { #3
$Continue = Read-host "- Do you want to continue processing $DisabledUser ? Y/N"
if ($continue -eq "Y" -or $continue -eq "y" ) { #4
write-host "processing $DisabledUser"
} #4
else { #4
write-host "Not processing $DisabledUser"
$notes = $false
}#4
}##
# if the information is recorded in notes, then start deleting stuff
if ($notes -eq $true ) { #3
# If there is a manager set, remove them
if ($DisabledUser.manager -ne $null ) {
    write-host "- Removing manager"
    $manager = Remove_Manager -user $DisabledUser
    $ListProcessedUsers += "," + $manager
}
#If there is no manager
else {
    $ListProcessedUsers += ","
    $manager = $true
}
# we've done this before, but it should be OK
#$getGroups = get-aduser $DisabledUser -properties memberof | select -expand memberof
#If there are groups
# if ( $getGroups -ne $null) {
if ( $notes.groups -ne $null) {
    write-host "- Removing email and security groups"
    # $email = Remove_EmailandSecurityGroups -user $DisabledUser $getGroups $Server
    $email = Remove_EmailandSecurityGroups -user $DisabledUser $notes.groups $Server
    $ListProcessedUsers += "," + $email
}
#Else if there are no groups
else {
    $ListProcessedUsers += ","
    # this is OK, because the email groups have been recorded
    $email = $true
}
#Hide from address book
write-host "- Hiding from address book"
$hideAddress = Hide_fromAddressBook -user $DisabledUser
$ListProcessedUsers += "," + $hideAddress
#Hide from Skype
write-host "- Hiding from Skype"
$hideSkype = Hide_fromSkype -user $DisabledUser
$ListProcessedUsers += "," + $hideSkype
} # 3
# not sure there is a point to doing this one:
#$ArchiveEmail = Archive_EmailAccount -user $DisabledUser
# not sure what to do with this part
# $HomeDrive = Process_HomeDrive -user $DisabledUser
#tested OK
# if these all worked, then mark the account as being processed
if ( $hideSkype -eq $true -and $manager -eq $true -and $email -eq $true -and $hideAddress -eq $true )
{ #3
write-host "Updating description to Processed"
$processed = Mark_AccountAsProcessed -user $DisabledUser
$ListProcessedUsers += "," + $processed + "`r`n"
}#3
} #2
} #1
# record a list of processed users, and thier info to a file, just in case.
try {
    $ListProcessedUsers | Out-file -FilePath "c:\temp\DeletedUsersProcessed.csv" -append
    write-host "writing to file"
}
catch {
    #error writing to file
    write-host "Unable to write list to file"
    #exit the script
}
return $ListProcessedUsers
}#0

function Write_UserInformationToNotes {
    [cmdletbinding()]
    param(
    [Parameter(Mandatory=$True)]
    [Microsoft.ActiveDirectory.Management.ADAccount]$User,
    [Parameter(Mandatory=$True)]
    [string]$Server)
    # get the group info
    write-host ""
    write-host "Reading group info from server" $Server
    #$groups = $null
    $getGroups = get-aduser $User -properties memberof | select -expand memberof
    $groups = read_GroupInfo $user $getGroups
    # get the manager info
    write-host ""
    write-host "Reading manager info"
    $manager = $null
    $manager = read_MgrInfo -user $user -server $Server
    # ignore it if the manager or the groups are empty
    try {
        write-host ""
        write-host "writing groups and manager to notes"
        set-aduser $user -replace @{info=$groups + $manager} @JustTesting
        # set-aduser $user -replace @{info=$groups + $manager} -whatif
        $ReturnValue = $true
    }
    catch {
        write-host "writing info to notes has failed!"
        # log error too
        $ReturnValue = $false
    }
    #could also return the manager and groups, since they are used again later....
    return $ReturnValue,$manager,$groups
}
function Read_UserInformationFromNotes {
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User,[string]$Server)
    # read info from Notes field, to make sure it was recorded
    try {
        $info = get-aduser $user -properties * | select info
        write-host $info
    }
    catch {
        write-host "Reading info to notes has failed!"
        $ReturnValue = $false
    }
    return $ReturnValue
}
function read_GroupInfo {
    # record existing email and security groups and manager
    param($User,$grouplist)
    # get the groups the user is in
    if ($grouplist) {
        write-host "Reading groups for "$user.name
        #blank the string for user, add this bit in:
        $groupString="Security and email groups:`r`n"
        # loop through each group
        foreach ( $group in $groupList ) {
            $MyGroup = Get-ADgroup $group -server $Server
            #if its not the "domain users" group
            if ( $MyGroup.name -ne "Domain Users" ) {
                #add the group to the string
                write-host "found group: "$MyGroup.name
                $groupString += $MyGroup.name + "`r`n"
            }
        }
    }
    # if there are no groups
    else {
        write-host "**No groups found"
        $groupString += "**No groups found`r`n"
    }
    #return this to update the note field with the email groups
    return $groupString
}
function read_MgrInfo {
    # record existing email and security groups and manager
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User,[string]$Server)
    $managerName = $null
    $mgrString = $null
    # get the name of the manager
    write-host "getting manager account for" $user.name
    if ($user.manager -ne $null ) {
        # record the manager name in the string
        $managerName = get-aduser $user.manager -properties * -server $Server
        write-host "Manager is " $managerName.Name
        $MgrString = "Manager: " + $managerName.name + "`r`n"
    }
    else {
        # can't find a manager
        write-host "**No manager found"
        $MgrString = "**No manager found"
    }
    #return this to update the note field with the manager
    return $MgrString
}
Function Remove_EmailandSecurityGroups {
    # Remove email and security groups
    param($User,$groupList,$DomainController)
    param(
    [Parameter(Mandatory=$True)]
    [string]$fromOU,
    [Parameter(Mandatory=$True)]
    [string]$Server,
    [Parameter(Mandatory=$True)]
    [string]$toOU)
    # get the list of groups
    # = get-aduser $user -properties memberof | select -expand memberof
    write-host "removing: " $user.name
    foreach ( $group in $groupList ) {
        # each of these is a hit, might be nicer to remove a bunch of people at once instead, but....
        try {
            $groupName = get-adobject $group -server $server
            write-host " from "$groupName.name
            # Uses AD to remove from group
            # Remove-ADGroupMember $groupname -members $user -Confirm:$false @JustTesting
            # Uses Exchange to remove from group
            Remove-DistributionGroupMember $groupname.name -member $user.name -Confirm:$false @JustTesting -DomainController $DomainController
            #
            Remove-DistributionGroupMember <distributionGroup> -member <MemberName> -Confirm:$false @JustTesting -DomainController <Dom2DomainController>
            #Remove-DistributionGroupMember <distributionGroup> -member <MemberName> -Confirm:$false @JustTesting -DomainController <Dom1DomainController>
            # Remove-ADGroupMember $groupname -members $user -Confirm:$false -whatif
            $Returnvalue = $true
        }
        catch
        {
            write-host $error[0]
            # if it breaks, return false
            $Returnvalue = $false
        }
    }
    # return a result
    return $ReturnValue
}
Function Remove_Manager {
    # Blank the "Manager" field
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # remove the manager info from the account
    try {
        set-aduser $user -manager $null
        # set-aduser $user -manager $null -whatif
        write-host "clearing manager info"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Hide_fromAddressBook {
    # Hide from Exchange address book
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # hide from GAL address book
    try {
        set-aduser -identity $user -replace @{msExchHideFromAddressLists=$true} @JustTesting
        # set-aduser -identity $user -replace @{msExchHideFromAddressLists=$true} -whatif
        write-host "hiding account from GAL"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Hide_fromSkype {
    # hide the account from Skype address book
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # hide from Skype
    try {
        set-adobject $user -clear showInAddressBook @JustTesting
        # set-adobject $user -clear showInAddressBook -whatif
        write-host "hiding account from Skype address book"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Process_HomeDrive {
    #Do something with the home drive
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # Not sure what to do here to be honest
    write-host "What do i do? I dont do anythng!"
    #Fake Value!
    $ReturnValue = $true
    return $ReturnValue
}
Function Archive_EmailAccount {
    # Archive the email account.
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    #load exchange modules
    #find email account
    #hmm, will this still work for measurement people/glycosyn etc?
    #$SMTPAddress = get-aduser $User -Properties * | select UserPrincipalName
    #get-mailbox
    # change it to archive
    #Fake Value!
    $ReturnValue = $true
    return $ReturnValue
}
Function Mark_AccountAsProcessed {
    # Mark account as processed
    #param([ADaccount]$User)
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    write-host "Setting user as processed:" $User.name
    # update the description field to say the account was processed
    try {
        $description = "Account Disabled: " + $User.description
        set-aduser $user -Description $description @JustTesting
        write-host "Success!"
        $ReturnValue = $true
    }
    catch {
        write-host "Fail :("
        $ReturnValue = $false
    }
    return $ReturnValue
}
Function Email_listToHelpdesk {
    # Email processed accounts to the helpdesk
    param([array]$Dom1UserList,[array]$Dom2UserList)
    # email the Service desk with the list of users which have been processed
    <#
    foreach ($user in $Dom1UserList ) {
    # do some stuff
    Write-host "I'm preparing to send an email"
    }
    foreach ($usr in $Dom2UserList ) {
    # do some stuff
    Write-host "I'm preparing to send an email too!"
    }
    #send an email (using which account?) to the helpdesk. hurrah!
    $MailArgs = @{
    From = $fromAddress
    To = $helpdeskAddress
    Subject = "Processing of departed users"
    # lets pretend this will work
    Body = $Dom2UserList, $Dom1UserList
    SmtpServer = 'smtp.office365.com'
    Port = 587
    UseSsl = $true
    Credential = $Credentials
    }
    Send-MailMessage @MailArgs -BodyAsHtml
    #>
}

# Connect to Exchange to remove from email groups
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://exch01.<ExchangeServerFQDN>/powershell/ -Credential $UserCredential -AllowRedirection -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck)
    Import-PSSession $Session -DisableNameChecking -AllowClobber
}
Catch {
    Write-Host "ERROR Connection to Exchange FAILED!" -ForegroundColor Red
    break
}
# returns a list of moved users in Dom1 domain:
Move_DisabledUsers -fromOU <DC1Ou1> -toOU <DC1Ou2> -Server <Dom1DomainController>
Move_DisabledUsers -fromOU <DC2Ou1> -toOU <DC2Ou2> -Server <Dom2DomainController>
# processes the users in the departed user's folder
# $Dom1UserList =
process_DisabledUsers -DepartedUsersOU <departedUsersOU> -Server <Dom1DomainController>
# $Dom2UserList =
process_DisabledUsers -DepartedUsersOU <departedUsersOU> -Server <Dom2DomainController>
#Email_listToHelpdesk -Dom1UserList $Dom1UserList -Dom2UserList $Dom2UserList
# disconnect from Exchange server:
remove-pssession $session


#Move disabled users from Users OU to Disabled users OU before we do anything else
[cmdletbinding()]
param(
[Parameter(Mandatory=$True)]
[string]$fromOU,
[Parameter(Mandatory=$True)]
[string]$Server,
[Parameter(Mandatory=$True)]
[string]$toOU)
#$usersToMove = $NULL
if ($fromOU -eq $null -or $Server -eq $null -or $toOU -eq $null)
{
    write-host "You need to add the data for function Move_DisabledUsers"
    $returnvalue = $false
}
else
{
    # find disabled users in the OU
    try {
        $usersToMove = get-aduser -searchbase $fromOU -searchscope 1 -filter { enabled -eq $false } -Properties * -Server $Server | sort name
    }
    catch {
        # if it doesn't work, or is empty (no users found) return nothing
        write-host "no users found to move."
        #$returnvalue = $false
    }
    if ($usersToMove) {
        $returnvalue=$null
        foreach ($MoveThisUser in $usersToMove )
        {
            write-host "moving: " $MoveThisUser.Samaccountname
            # check for errors here
            try {
                move-adobject -identity $MoveThisUser -server $Server -targetpath $toOU @JustTesting
                # write out a list of the moved users
                #record the moved users and return to the script
                #$returnvalue += $usersToMove.name
            }
            catch {
                write-host "unable to move user:" $error[0]
                # write an error instead to the list of moved users
                #$returnvalue += $usersToMove.name +
            }
        }
    }
}
# return $returnvalue
}
function process_DisabledUsers {
    # check if the account has already been processed, if not, then process it
    [cmdletbinding()]
    param(
    [Parameter(Mandatory=$True)]
    [string]$DepartedUsersOU,
    [Parameter(Mandatory=$True)]$Dom2UserList
    [string]$Server)
    write-host $DepartedUsersOU
    $DisabledUserList = get-aduser -searchbase $DepartedUsersOU -filter { enabled -eq $false } -Properties * -Server $Server | sort name
    # change for testing, this just returns one user rathert than a huge list.
    # $DisabledUserList = get-aduser -filter { enabled -eq $false -and userPrincipalName -eq <sampleuser> } -Properties * -Server $Server
    # use a string to store the results
    $ListProcessedUsers = "Name, info written, manager remove, email hide, skype hide, processed `r`n"
    foreach ( $DisabledUser in $DisabledUserList ) {#1
    write-host " "
    write-host "### Processing " $disableduser.name
    # if Disabled user already processed (ie description starts with "Disabled"
    if ( $DisabledUser.description.contains("Disabled") ) {#2
    write-host "user " $disableduser.name " has already been processed"
}#2
# error reading the notes?
# otherwise:
else { #2
write-host "user " $disableduser.name " has not been processed, processing"
$ListProcessedUsers += $disableduser.name
# reset the variables
$hideSkype = $false
$manager = $false
$email = $false
$hideAddress = $false
# info out into notes make sure it is recorded
write-host "writing information to Notes field from "$Server
$notes = Write_UserInformationToNotes -user $DisabledUser -Server $Server
# read the info out to make sure it is recorded.
if ( Read_UserInformationFromNotes -user $DisabledUser -Server $Server )
{
    $ListProcessedUsers += "," + $notes
}
# prompt to continue if this is a test:
if ($prompt) { #3
$Continue = Read-host "- Do you want to continue processing $DisabledUser ? Y/N"
if ($continue -eq "Y" -or $continue -eq "y" ) { #4
write-host "processing $DisabledUser"
} #4
else { #4
write-host "Not processing $DisabledUser"
$notes = $false
}#4
}##
# if the information is recorded in notes, then start deleting stuff
if ($notes -eq $true ) { #3
# If there is a manager set, remove them
if ($DisabledUser.manager -ne $null ) {
    write-host "- Removing manager"
    $manager = Remove_Manager -user $DisabledUser
    $ListProcessedUsers += "," + $manager
}
#If there is no manager
else {
    $ListProcessedUsers += ","
    $manager = $true
}
# we've done this before, but it should be OK
#$getGroups = get-aduser $DisabledUser -properties memberof | select -expand memberof
#If there are groups
# if ( $getGroups -ne $null) {
if ( $notes.groups -ne $null) {
    write-host "- Removing email and security groups"
    # $email = Remove_EmailandSecurityGroups -user $DisabledUser $getGroups $Server
    $email = Remove_EmailandSecurityGroups -user $DisabledUser $notes.groups $Server
    $ListProcessedUsers += "," + $email
}
#Else if there are no groups
else {
    $ListProcessedUsers += ","
    # this is OK, because the email groups have been recorded
    $email = $true
}
#Hide from address book
write-host "- Hiding from address book"
$hideAddress = Hide_fromAddressBook -user $DisabledUser
$ListProcessedUsers += "," + $hideAddress
#Hide from Skype
write-host "- Hiding from Skype"
$hideSkype = Hide_fromSkype -user $DisabledUser
$ListProcessedUsers += "," + $hideSkype
} # 3
# not sure there is a point to doing this one:
#$ArchiveEmail = Archive_EmailAccount -user $DisabledUser
# not sure what to do with this part
# $HomeDrive = Process_HomeDrive -user $DisabledUser
#tested OK
# if these all worked, then mark the account as being processed
if ( $hideSkype -eq $true -and $manager -eq $true -and $email -eq $true -and $hideAddress -eq $true )
{ #3
write-host "Updating description to Processed"
$processed = Mark_AccountAsProcessed -user $DisabledUser
$ListProcessedUsers += "," + $processed + "`r`n"
}#3
} #2
} #1
# record a list of processed users, and their info to a file, just in case.
try {
    $ListProcessedUsers | Out-file -FilePath "c:\temp\DeletedUsersProcessed.csv" -append
    write-host "writing to file"
}
catch {
    #error writing to file
    write-host "Unable to write list to file"
    #exit the script
}
return $ListProcessedUsers
}#0

function Write_UserInformationToNotes {
    [cmdletbinding()]
    param(
    [Parameter(Mandatory=$True)]
    [Microsoft.ActiveDirectory.Management.ADAccount]$User,
    [Parameter(Mandatory=$True)]
    [string]$Server)
    # get the group info
    write-host ""
    write-host "Reading group info from server" $Server
    #$groups = $null
    $getGroups = get-aduser $User -properties memberof | select -expand memberof
    $groups = read_GroupInfo $user $getGroups
    # get the manager info
    write-host ""
    write-host "Reading manager info"
    $manager = $null
    $manager = read_MgrInfo -user $user -server $Server
    # ignore it if the manager or the groups are empty
    try {
        write-host ""
        write-host "writing groups and manager to notes"
        set-aduser $user -replace @{info=$groups + $manager} @JustTesting
        # set-aduser $user -replace @{info=$groups + $manager} -whatif
        $ReturnValue = $true
    }
    catch {
        write-host "writing info to notes has failed!"
        # log error too
        $ReturnValue = $false
    }
    #could also return the manager and groups, since they are used again later....
    return $ReturnValue,$manager,$groups
}
function Read_UserInformationFromNotes {
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User,[string]$Server)
    # read info from Notes field, to make sure it was recorded
    try {
        $info = get-aduser $user -properties * | select info
        write-host $info
    }
    catch {
        write-host "Reading info to notes has failed!"
        $ReturnValue = $false
    }
    return $ReturnValue
}
function read_GroupInfo {
    # record existing email and security groups and manager
    param($User,$grouplist)
    # get the groups the user is in
    if ($grouplist) {
        write-host "Reading groups for "$user.name
        #blank the string for user, add this bit in:
        $groupString="Security and email groups:`r`n"
        # loop through each group
        foreach ( $group in $groupList ) {
            $MyGroup = Get-ADgroup $group -server $Server
            #if its not the "domain users" group
            if ( $MyGroup.name -ne "Domain Users" ) {
                #add the group to the string
                write-host "found group: "$MyGroup.name
                $groupString += $MyGroup.name + "`r`n"
            }
        }
    }
    # if there are no groups
    else {
        write-host "**No groups found"
        $groupString += "**No groups found`r`n"
    }
    #return this to update the note field with the email groups
    return $groupString
}
function read_MgrInfo {
    # record existing email and security groups and manager
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User,[string]$Server)
    $managerName = $null
    $mgrString = $null
    # get the name of the manager
    write-host "getting manager account for" $user.name
    if ($user.manager -ne $null ) {
        # record the manager name in the string
        $managerName = get-aduser $user.manager -properties * -server $Server
        write-host "Manager is " $managerName.Name
        $MgrString = "Manager: " + $managerName.name + "`r`n"
    }
    else {
        # can't find a manager
        write-host "**No manager found"
        $MgrString = "**No manager found"
    }
    #return this to update the note field with the manager
    return $MgrString
}
Function Remove_EmailandSecurityGroups {
    # Remove email and security groups
    param($User,$groupList,$DomainController)
    param(
    [Parameter(Mandatory=$True)]
    [string]$fromOU,
    [Parameter(Mandatory=$True)]
    [string]$Server,
    [Parameter(Mandatory=$True)]
    [string]$toOU)
    # get the list of groups
    # = get-aduser $user -properties memberof | select -expand memberof
    write-host "removing: " $user.name
    foreach ( $group in $groupList ) {
        # each of these is a hit, might be nicer to remove a bunch of people at once instead, but....
        try {
            $groupName = get-adobject $group -server $server
            write-host " from "$groupName.name
            # Uses AD to remove from group
            # Remove-ADGroupMember $groupname -members $user -Confirm:$false @JustTesting
            # Uses Exchange to remove from group
            Remove-DistributionGroupMember $groupname.name -member $user.name -Confirm:$false @JustTesting -DomainController $DomainController
            # Remove-ADGroupMember $groupname -members $user -Confirm:$false -whatif
            $Returnvalue = $true
        }
        catch
        {
            write-host $error[0]
            # if it breaks, return false
            $Returnvalue = $false
        }
    }
    # return a result
    return $ReturnValue
}
Function Remove_Manager {
    # Blank the "Manager" field
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # remove the manager info from the account
    try {
        set-aduser $user -manager $null
        # set-aduser $user -manager $null -whatif
        write-host "clearing manager info"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Hide_fromAddressBook {
    # Hide from Exchange address book
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # hide from GAL address book
    try {
        set-aduser -identity $user -replace @{msExchHideFromAddressLists=$true} @JustTesting
        # set-aduser -identity $user -replace @{msExchHideFromAddressLists=$true} -whatif
        write-host "hiding account from GAL"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Hide_fromSkype {
    # hide the account from Skype address book
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # hide from Skype
    try {
        set-adobject $user -clear showInAddressBook @JustTesting
        # set-adobject $user -clear showInAddressBook -whatif
        write-host "hiding account from Skype address book"
        $ReturnValue = $True
    }
    catch {
        $ReturnValue = $false
    }
    # return a result
    return $ReturnValue
}
Function Process_HomeDrive {
    #Do something with the home drive
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    # Not sure what to do here to be honest
    write-host "What do i do? I dont do anythng!"
    #Fake Value!
    $ReturnValue = $true
    return $ReturnValue
}
Function Archive_EmailAccount {
    # Archive the email account.
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    #load exchange modules
    #find email account
    #$SMTPAddress = get-aduser $User -Properties * | select UserPrincipalName
    #get-mailbox
    # change it to archive
    #Fake Value!
    $ReturnValue = $true
    return $ReturnValue
}
Function Mark_AccountAsProcessed {
    # Mark account as processed
    #param([ADaccount]$User)
    param([Microsoft.ActiveDirectory.Management.ADAccount]$User)
    write-host "Setting user as processed:" $User.name
    # update the description field to say the account was processed
    try {
        $description = "Account Disabled: " + $User.description
        set-aduser $user -Description $description @JustTesting
        write-host "Success!"
        $ReturnValue = $true
    }
    catch {
        write-host "Fail :("
        $ReturnValue = $false
    }
    return $ReturnValue
}
Function Email_listToHelpdesk {
    # Email processed accounts to the helpdesk
    param([array]$Dom1UserList,[array]$Dom2UserList)
    # email the Service desk with the list of users which have been processed
    <#
    foreach ($user in $Dom1UserList ) {
    # do some stuff
    Write-host "I'm preparing to send an email"
    }
    foreach ($usr in $Dom2UserList ) {
    # do some stuff
    Write-host "I'm preparing to send an email too!"
    }
    #send an email (using which account?) to the helpdesk. hurrah!
    $MailArgs = @{
    From = <senderAddress>
    To = <helpdeskAddress>
    Subject = "Processing of departed users"
    # lets pretend this will work
    Body = $Dom2UserList, $Dom1UserList
    SmtpServer = 'smtp.office365.com'
    Port = 587
    UseSsl = $true
    Credential = $Credentials
    }
    Send-MailMessage @MailArgs -BodyAsHtml
    #>
}

# Connect to Exchange to remove from email groups
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://<exchangeServerFQDN>/powershell/ -Credential $UserCredential -AllowRedirection -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck)
    Import-PSSession $Session -DisableNameChecking -AllowClobber
}
Catch {
    Write-Host "ERROR Connection to Exchange FAILED!" -ForegroundColor Red
    break
}
# returns a list of moved users in Dom1 domain:
Move_DisabledUsers -fromOU <Dom1OU> -toOU <Dom2OU> -Server <Dom1DC>
# processes the users in the departed user's folder
# $Dom1UserList =
process_DisabledUsers -DepartedUsersOU <Ou1> -Server <Dom1DomainController>
# $Dom2UserList =
process_DisabledUsers -DepartedUsersOU <Ou2> -Server  <Dom2DomainController>
#Email_listToHelpdesk -Dom1UserList $Dom1UserList -Dom2UserList $Dom2UserList
# disconnect from Exchange server:
remove-pssession $session

