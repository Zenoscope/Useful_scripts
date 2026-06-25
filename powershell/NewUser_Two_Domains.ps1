##### A script to create a new user.###
<#
.SYNOPSIS
Create a new user in both Dom1 and Dom2 domains.
.DESCRIPTION
Creates an AD account in Dom2 and Dom1, an exchange mailbox (and adds to the site specific mail group too)
.NOTES
Author: Ryan McArthur
Last Edit: 2019-08-20
Version 1.0 - initial release
Todo:
- Add to mail groups, add mail account
- Put in diff OU based on office location/group?
- status box with current process being performed, OK runs script, not closes window.
- add (hidden) tick box for women
#>
<#
Info needed to create an account
$firstname = first name
$LastName = last name
$PhysOffice = Physical Office location
$Title = Position (Job Title)
$Dept = "Unit"
$Group = Group"
$Team = Team"
$Magr = Manager's UserName"
$today = get-date -UFormat "%d/%m/%Y"
$endDate = Read-Host "Please enter end date for account, eg $today (optional)"
$Password = Read-Host "ENTER Password"
Steps:
Create the account in both AD domains will all information required
Add to the default groups
Create the account in Exchange
Add the home drive
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
# clear the error message
$error.clear()
#----------------------------------------------------------[Declarations]----------------------------------------------------------
# sets the domain controller where the accounts are first made/manipulated.
$ExchangeServer = "exch02.Dom1.govt.nz" #exchange server
$Dom1_DC   = "dc01.Dom1.govt.nz" # PDC for Dom1 domain
$Dom2_DC = "grdns2.int.irl.cri.nz" # PDC for Dom2 domain
$db = "MBDB01", "MBDB02", "MBDB03", "MBDB04", "MBDB04", "MBDB05", "MBDB06", "MBDB07", "MBDB08" # exchange databases
# get the one which is smallest?
$MailboxDB  = Get-Random $db # pick a random database
$Dom1OU = "OU=Sync Inbound,OU=Dom1 People,OU=People,DC=Dom1,DC=govt,DC=nz" # OU for user accounts
$Dom2OU = "OU=Sync Outbound,OU=Dom1 User Accounts,DC=int,DC=irl,DC=cri,DC=NZ" # OU for user accounts
##
# test OU for users
#$Dom1OU = "OU=Test,OU=Dom1 People,OU=People,DC=Dom1,DC=govt,DC=nz"
#$Dom2OU = "OU=Test Users,OU=Sync Outbound,OU=Dom1 User Accounts,DC=int,DC=irl,DC=cri,DC=nz"
#
$Dom1domain = "Dom1.govt.nz"
$Dom2domain = "int.irl.cri.nz"
$Dom2CN = "Dom2"
$Dom1CN = "Dom1"
#$ErrorActionPreference = "Stop"
$Date = Get-Date
$HomeDrivePath = "\\fs01\home\Users\"
# addresses
$AddressTable = @{
    # North Harbour
    1 = @{
        StreetAddress = "100 Harbour View Drive"
        POBox = "PO Box 10001"
        City = "North Harbour"
        State = "North Island"
        St = "North Island"
        PostalCode = "1001"
        Office = "North Harbour"
    }

    # Riverside
    2 = @{
        StreetAddress = "250 Riverside Parkway"
        POBox = "PO Box 20002"
        City = "Riverside"
        State = "Central Region"
        St = "Central Region"
        PostalCode = "2002"
        Office = "Riverside"
    }

    # Southern Cross
    3 = @{
        StreetAddress = "75 Southern Cross Avenue"
        POBox = "PO Box 30003"
        City = "Southern Cross"
        State = "South Region"
        St = "South Region"
        PostalCode = "3003"
        Office = "Southern Cross"
    }

    # Metro Central
    4 = @{
        StreetAddress = "500 Innovation Boulevard"
        POBox = "PO Box 40004"
        City = "Metro Central"
        State = "Northern Region"
        St = "Northern Region"
        PostalCode = "4004"
        Office = "Metro Central"
    }
}
}
#-----------------------Function declarations-----------------------

# this function adds the text labels into the GUI
function add-label {
    # these are the function's arguments
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [int]$XLoc,
    [Parameter (Mandatory=$true)]
    [int]$YLoc,
    [Parameter (Mandatory=$false)]
    [int]$width
    )
    $Label = New-Object System.Windows.Forms.Label
    # put it somewhere
    $Label.Location = New-Object System.Drawing.Size($XLoc,$YLoc)
    # this is the text of the label
    $Label.Text = $name
    # sets the width of the text label, if it is unset
    if ($width -eq "") { $width = 80 }
    $Label.Size = New-Object Drawing.Size($width,20)
    #return the object to use it
    $winform.controls.add($Label)
    return $label
}
# Add an alert box
function add-AlertDialog {
    # these are the function's arguments
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [string]$title,
    [Parameter (Mandatory=$true)]
    [string]$message
    )
    #Also pop up a window or something, go back to not create the new user.
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    # display an "OK" button only
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    # set the icon as "error"
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    # set the title and message for the dialog box.
    $MessageBody = $message
    $MessageTitle = $title
    $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
}
# add a text box
function add-Textbox {
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [int]$XLoc,
    [Parameter (Mandatory=$true)]
    [int]$YLoc
    )
    $Tbox = new-object System.Windows.Forms.Textbox
    $Tbox.Location = New-Object System.Drawing.Size($xloc,$yloc)
    $Tbox.Size = New-Object System.Drawing.Size(120,10)
    #write-host "making text box"
    $winform.controls.add($Tbox)
    return $Tbox
}
# add a button
function add-Button {
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [string]$name,
    [Parameter (Mandatory=$true)]
    [int]$XLoc,
    [Parameter (Mandatory=$true)]
    [int]$YLoc,
    )
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Point($xloc,$yloc)
    $Button.Size = New-Object System.Drawing.Size(100,23)
    $Button.Text = $name
    $Button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $winform.controls.add($Button)
    return $Button
}
# add drop-down box
function add-Dropdown {
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [array]$droplist,
    [Parameter (Mandatory=$true)]
    [int]$xloc,
    [Parameter (Mandatory=$true)]
    [int]$yloc,
    [Parameter (Mandatory=$true)]
    [int]$dropheight,
    ########### "from" not used?
    [Parameter (Mandatory=$true)]
    [string]$from,
    #############
    [Parameter (Mandatory=$true)]
    [int]$width
    )
    $Drop = new-object System.Windows.Forms.Combobox
    $Drop.Location = New-Object System.Drawing.Size($xloc,$yloc)#
    $Drop.Size = New-Object Drawing.Size($width,20)
    #populate the values
    $count = 0
    ForEach($i in $DropList)
    { $Drop.items.add($i) | Out-Null
    $count++ | Out-Null
}
#sets the height of the drop down
$Drop.DropDownHeight=$dropheight
$Drop.add_SelectedIndexChanged($script:select)
$winform.controls.add($Drop)
return $drop
}
# sets a temporary password for the account
Function Get-Temppassword() {
    [cmdletbinding()]
    Param(
    [Parameter (Mandatory=$true)]
    [int]$length=10
    )
    $sourcedata = $NULL;For ($a=33;$a –le 126;$a++) {$ascii+=,[char][byte]$a }
    For ($loop=1; $loop –le $length; $loop++) {
        $TempPassword += ($sourcedata | GET-RANDOM)
    }
    return $TempPassword
}
# change the folder ACL for the home folder
function Add-Access() {
    [cmdletbinding()]
    param(
    [Parameter (Mandatory=$true)]
    [string]$username,
    [string]$folder
    )
    # set the access rule (permissions) on the folder
    $NewAccessrule = New-Object System.Security.AccessControl.FileSystemAccessRule $username, 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow'
    # get the existing permissions
    $currentACL = Get-Acl -Path $folder
    # add them on
    $currentACL.AddAccessRule($NewAccessrule)
    # set the SCL on the folder
    Set-Acl -Path $folder -AclObject $currentACL
    #check to see if it worked:
    # get the permissions
    $Access_String_Array = get-Acl -Path $userfolder | select accesstostring
    # search for the permissions and generate an error if they aren't there:
    #if ( ( $Access_String_Array.contains("$userName")) -eq $false ) {
    # add-AlertDialog -title "Error setting home folder permissons" , -message "Folder permissions not set on $folder for $userName"
    # }
}
# looks like a function, functions like a variable....
# finds the manager in AD
$Manager_Select = {
    $manager_name = ""
    write-host "updating Manager box"
    #gets the $manager.text variable to use in the script
    #$script:
    $manager_find = $manager.text
    write-host "updating Manager to: $manager_find"
    # search for the managers name, error alert if not found.
    # use 2> $NULL to supress the error because it will always be blank the first time it is searched, though we could skip this if it was blank, or the first search, or make it search when focus leaves the box instead.
    #try {
    $manager_name = get-aduser -properties * -filter { name -like $manager_find } -SearchBase "OU=Dom1 People,OU=People,DC=company1,DC=org" -Credential $credential 2> $NULL
    write-host $manager_name.samaccountname
    # }
    # if it breaks, then its bad.
    # if $error[0] {}
    if ($manager_name) {
        write-host "found manager" $manager_name.samaccountname
    }
    else {
        write-host "Manager doesn't exist!"
        add-AlertDialog -title "Manager" -message "The manager you entered for was not found. Check your spelling!"
    }
    $script:manager_text=$manager_name.samaccountname
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
try
{
    #Import Active Directory Module
    Import-Module ActiveDirectory
}
Catch
{
    Write-Host "ERROR Connection to AD Module FAILED!" -ForegroundColor Red
    throw
}
# ask for an admin password, maybe a while loop instead?
if ( -Not $Credential ) {
    #only ask if it hasn't already been set
    $Credential = Get-Credential $env:USERNAME
}
#
#--------------------- Draw the form
#
##-------draw the form itself:
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | out-null
$WinForm = new-object Windows.Forms.Form
$WinForm.TopMost = $True
# set the name of the window:
$WinForm.text = "New User Creation"
#set the size of the window: Y by X
$WinForm.Size = new-object Drawing.Size(750,600)
#$WinForm.AutoSize = $True
$WinForm.AutoSizeMode = "GrowAndShrink"
$Font = New-Object System.Drawing.Font("Times New Roman",11,[System.Drawing.FontStyle]::Regular)
# Font styles are: Regular, Bold, Italic, Underline, Strikeout
$WinForm.Font = $Font
##-------Add labels
#they can all be called "label" because they don't need to return any value and thus be unique
$Label = add-label -name "First name:" -XLoc 20 -YLoc 15
$FirstName = add-Textbox -xloc 100 -yloc 10
# Last name
$Label = add-label -name "Last name:" -XLoc 250 -YLoc 15
$LastName = add-Textbox -xloc 330 -yloc 10
# job title
$Label = add-label -name "Position (Job title):" -XLoc 20 -YLoc 50 -width 150
$Title = add-Textbox -XLoc 200 -YLoc 50
# department
$Label = add-label -name "Unit (Department):" -XLoc 20 -YLoc 75 -width 150
$Dept = add-Textbox -xloc 200 -yloc 75
# Group
$Label = add-label -name "Group:" -XLoc 20 -YLoc 100
$Group = add-Textbox -xloc 200 -yloc 100
# team
$Label = add-label -name "Team:" -XLoc 20 -YLoc 125
$Team = add-Textbox -xloc 200 -yloc 125
# Manager
$Label = add-label -name "Reports to (Manager):" -XLoc 20 -YLoc 150
$manager = add-Textbox -xloc 200 -yloc 150
$manager.add_lostfocus($Manager_select)
# Office
$Label = add-label -name "Location (Office):" -XLoc 20 -YLoc 175
$PhysOffice_list = @("North Harbour", "Riverside", "South Region","Northern Region")
$PhysOffice = add-Dropdown -droplist $PhysOffice_list -xloc 200 -yloc 175 -dropheight 200 -width 230 -from 'PhysOffice'
# end date
$Label = add-label -name "End Date:" -XLoc 20 -YLoc 200
$Label = add-label -name "D:" -XLoc 180 -YLoc 205 -width 25
$Label = add-label -name "M:" -XLoc 250 -YLoc 205 -width 25
$Label = add-label -name "Y:" -XLoc 340 -YLoc 205 -width 25
##-------Add drop downs for end date
# day
$day_list=@()
for ($i = 1;$i -le 31;$i++) {
    #$thisMonth=
    $day_list += $i
}
$Day = add-Dropdown -droplist $day_list -xloc 205 -yloc 200 -dropheight 200 -width 40 -from 'Day'
# month
$month_list = @()
for ($i = 1;$i -le 12;$i++) {
    #$thisMonth=
    $month_list += get-date -UFormat %b -Month $i
}
$Month = add-Dropdown -droplist $month_list -xloc 275 -yloc 200 -dropheight 200 -width 50 -from 'Day'
# year
$this_year = get-date -UFormat %Y
$year_list=@()
for ($i = -1;$i -le 2;$i++) {
    $Year_list += [int]$this_year+$i
}
$Year=add-Dropdown -droplist $year_list -xloc 365 -yloc 200 -dropheight 100 -width 60 -from 'Day'
# extra interface options go here
# female tick box
# email for CRM
# possibly other groups they are automatically added to
##------- OK/cancel buttons
$OKButton = add-Button -name "Create User" -xloc 15 -yloc 420
$CancelButton = add-Button -name "Cancel" -xloc 15 -yloc 450
#-------- Display the form:
$WinForm.Add_Shown($WinForm.Activate())
$result = $WinForm.ShowDialog()
write-host $result
#-------- Buttons do this:
if ($result -eq "Cancel") {
    Write-Host -ForegroundColor Green "Tidying Up"
    Remove-Module ActiveDirectory
    Exit
}
if ($result -eq "Ok")
{
    write-host $result
    write-host "firstname is $firstname"
    write-host "firstname is $lastname"
}
#
#-------- Execute the form
#
# users need a manager, and first and last names, and a job title
if ( $FirstName.text -eq "" -and $result -ne "Cancel" ) { add-AlertDialog -title "First name cannot be blank" -message "You can't create a new user without a first name! WTF dude!"
throw
}
if ( $LastName.text -eq "" -and $result -ne "Cancel" ) { add-AlertDialog -title "Last name cannot be blank" -message "You can't create a new user without a last name! Sheesh!"
throw
}
if ( $title.text -eq "" -and $result -ne "Cancel" ) { add-AlertDialog -title "Last name cannot be blank" -message "You can't create a new user without a title, its unethical."
throw
}
if ( $Manager.text -eq "" -and $result -ne "Cancel" ) { add-AlertDialog -title "Last name cannot be blank" -message "You can't create a new user without a manager, they would just run wild."
throw
}
#set up the end date if there is one
if ($day.SelectedItem) {
    $day_ = [string]$day.SelectedItem
    write-host $day_
}
if($month.SelectedItem) {
    $month_ = [string]$month.SelectedItem
    write-host $month_
}
if ($year.SelectedItem ) {
    $year_ = [string]$year.SelectedItem
    write-host $year_
}
if ( $year_ -and $month -and $day_ ) {
    $endDate = "$day_/$month_/$year_"
    $EndDate_ = $endDate | get-date
    # acount expires at midnight that day, so add another day to correct it.
    $EndDateAD = $EndDate_.AddDays(1)
}
#
#------------------ Active Directory : Create the account
#
# create the account
# stuff you need:
# first name, last name, displayname, UPN, OU, password, change password at login, physicaloffice, manager, address, title/description, department, company, default group memberships
# -------generate temporary password
# generate a temp password
# https://blogs.technet.microsoft.com/heyscriptingguy/2013/06/03/generating-a-new-password-with-windows-powershell/
$Password = GET-Temppassword –length 20
# do this because the strings don't like dot syntax much.
$first_name_text = $firstname.text
$last_name_text = $lastname.text
$function_text = $Function.Text
$Office_text = $PhysOffice.SelectedItem
$Converted_Password = ConvertTo-SecureString -string $password -asplaintext -force
$email = "$first_name_text.$last_name_text@Dom1.org"
$sAMAcctName = $($first_name_text[0]+"."+$last_name_text)
$Dom2UPN = $first_name_text[0]+"."+$last_name_text+"@$Dom2domain"
$Dom1UPN = "$first_name_text.$last_name_text@$Dom1domain"
# formatting the description
if ($Team.text) { $Team_ = " - " + $Team.text }
if ($endDate) { $enddate_ = " - " + $endDate }
$description = $Title.text + $Team_ + $endDate_
# use this block for debugging the AD user creation.
write-host first: $first_name_text
write-host last: $last_name_text
write-host display name: "$first_name_text $last_name_text"
write-host name : "$first_name_text $last_name_text"
write-host SAM : $sAMAcctName
write-host "email: $email"
write-host manager: $manager_text
write-host Company: $Company_text
write-host dept: $Dept.text
write-host title: $title.Text
write-host desc: $description
write-host Dom1 UPN: $Dom1upn
write-host Dom1 OU: $Dom1OU
write-host Dom2 UPN: $Dom2UPN
write-host StreetAddress = $AddressTable.$site.StreetAddress
write-host POBox = $AddressTable.$site.POBox
write-host City = $AddressTable.$site.City
write-host State = $AddressTable.$site.State
write-host St = $AddressTable.$site.St
write-host PostCode = $AddressTable.$site.PostCode
write-host Office = $AddressTable.$site.Office
write-host End Date: $endDate
ping $Dom2_DC

# check to see if the user already exists, if they do, I dunno, something
If ( ( Get-ADUser -Server $Dom1_DC -Filter {sAMAccountname -eq $($first_name_text[0]+"."+$last_name_text) }) -or ( Get-ADUser -Server $Dom2_DC -Filter {sAMAccountname -eq $($first_name_text[0]+"."+$last_name_text)} ) ){
    add-AlertDialog -title "User exists!"-message "User already exists!"
    throw
}
$site = $PhysOffice.SelectedIndex
$NewUSerArguments = @{
    GivenName = $($firstname.text)
    Surname = $($lastname.text)
    DisplayName = "$($firstname.text) $($lastname.text)"
    Name = "$($firstname.text) $($lastname.text)"
    SamAccountName = $($first_name_text[0]+"."+$last_name_text)
    Email = $("$first_name_text.$last_name_text@Dom1.org")
    Department = $($Dept.text)
    Description = $Title.text + $Team_ + $endDate_
    Title = $title.Text
    AccountPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
    enabled = $true
    StreetAddress = $AddressTable.$site.StreetAddress
    POBox = $AddressTable.$site.POBox
    City = $AddressTable.$site.City
    State = $AddressTable.$site.State
    PostalCode = $AddressTable.$site.PostCode
    Office = $AddressTable.$site.Office
}
if ($EndDateAD) {
    $set_endDate = @{ AccountExpirationDate = $EndDateAD }
}
# if the end date is not set, don't pass anything at all
else {
    $set_endDate = @{}
}
try
# create the account in the Dom1 domain.
{
    #get-adorganizationalunit $Dom1OU
    New-ADUser -Server $Dom1_DC -Manager $manager_text -userprincipalname $Dom1upn -Path $Dom1OU -CannotChangePassword $true -PasswordNeverExpires $true @NewUSerArguments -Credential $Credential @set_endDate
}
catch
{
    Write-Host "Unknown error occured while creating account in $Dom1domain. Please note password entered must contact atleat one capital and number to meet the requirements of the domain" -ForegroundColor Red
    add-AlertDialog -title "Unable to create Dom1 account"-message "$error[0]"
    throw
}
try
# create the account in the Dom2 domain.
{
    New-ADUser -Server $Dom2_DC -userprincipalname $Dom2upn -Path $Dom2OU -ChangePasswordAtLogon $true @NewUSerArguments -Credential $Credential @set_endDate
}
catch
{
    Write-Host "Unknown error occured while creating account in $Dom2domain. Please note password entered must contact atleat one capital and number to meet the requirements of the domain" -ForegroundColor Red
    add-AlertDialog -title "Unable to create Dom2 account"-message "$error[0]"
    throw
}
if (get-ADUser -Server $Dom1_DC -Filter {sAMAccountname -eq $sAMAcctName} ) { write-host "User Created in Dom1 domain" }
if (get-ADUser -Server $Dom2_DC -Filter {sAMAccountname -eq $sAMAcctName} ) { write-host "User Created Dom2 domain" }
## Use this
#
# add to the default groups
# these should be dynamic groups one day....
#
try {
    add-adgroupmember -members $sAMAcctName -identity "Default Group" -server $Dom1_DC -credential $Credential
    #-whatif
}
catch { write-host "unable to add to group Default Group" -ForegroundColor Red }
try {
    add-adgroupmember -members $sAMAcctName -identity "SG_SP13_EDI_AllUsers" -server $Dom1_DC -credential $Credential
    #-whatif
}
catch { write-host "unable to add to group Default Group" -ForegroundColor Red }
#
#Create Home Drive
#
try{
    $userfolder = "\\fileservername\home\Users\$sAMAcctName"
    Write-Host "Creating Home Drive"
    if(!(Test-Path -Path $userfolder )){
        write-host "Creating folder $userfolder"
        New-Item -ItemType directory -Path $userfolder
    }
}
catch
{
    Write-Host "failed to create Home Drive" -ForegroundColor Red
}
#adds access and tests it too:
Add-Access -username "$Dom2domain\$sAMAcctName" -folder $userfolder
Add-Access -username "$Dom1domain\$($first_name_text+"."+$last_name_text)" -folder $userfolder
#
# ----------------------Create user and enable mailbox in Exchange
#
#Import session information for Exchange
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://exchangeserver.Dom1.org/powershell/ -Credential $Credential -AllowRedirection -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck)
    Import-PSSession $Session -DisableNameChecking -allowclobber
}
Catch {
    Write-Host "ERROR Connection to Exchange FAILED!" -ForegroundColor Red
    add-AlertDialog -title "Unable to connect to Exchange"-message "$error[0]"
    throw
}
# add the mailbox to the account
try {
    Enable-Mailbox -Displayname "$first_name_text $last_name_text" -database $MailboxDB -Alias "$first_name_text.$last_name_text" -domainController $Dom1_DC -identity $sAMAcctName
    Start-Sleep -Seconds 25
}
catch {
    Write-Host "An error occured while creating Exchange account." -ForegroundColor Red
    add-AlertDialog -title "An error occured while creating Exchange account."-message "$error[0]"
    #Exit
    throw
}
# pause here while the account is being created
#If there is a site DL (*Site-Staff-<location>) then add to the site DL
try {
    $Office_Location= "*Staff-"+$Office_text+"*"
    if ( get-ADobject -filter { name -like $Office_Location } | add-adgroupmember -members $sAMAcctName ) {
        write-host "Added to $Office_Location"
    }
}
catch {
    write-host "Unable to add user to group: *Site-Staff-$Office_text"
}
# Add to the Woman DL if they are permanent:
if ($woman -eq $true) {
    switch ( $Office_text ) {
        "Asteron" { $Area = "Wellington" }
        "Gracefield" { $Area = "Wellington" }
        "Auckland" { $Area = "Auckland" }
        "Christchurch" { $Area = "Christchurch" }
    }
    try {
        $Women_Location= "*_Woman-"+$Area
        get-ADobject -filter { name -like $Women_Location } | add-adgroupmember -members $sAMAcctName
        Write-host "Added to: "$Women_Location
    }
    catch {
        Write-host "unable to add to Women's email list"
    }
}
<#
# if user is CRM user, then send an email:
# check if ADM accounts have send-as rights, or email accounts
$Email_Message = @{
To = "CrmHelp@Dom1.org"
From = "helpdesk@Dom1.org"
Subject "New CRM user"
Body = "Hi,<br> Can you please setup $FirstName $LastName as a CRM user.<br> Thanks,<br> Digital Service Desk"
BodyAsHtml = $true
}
Send-MailMessage @Email_Message
$Email_Message = @{
To = $("$first_name_text.$last_name_text@Dom1.org")
From = "helpdesk@Dom1.org"
Subject "New CRM user"
Body = "Hi,<br> Can you please setup $FirstName $LastName as a CRM user.<br> Thanks,<br> Digital Service Desk"
BodyAsHtml = $true
}
Send-MailMessage @Email_Message
#>

################################################### Maybe later...
<#
#f there is a team DL (*Staff-<Group>-<Unit>)
##don’t use the “manual” one
try {
# unit = dept
# group = group
if ($Dept.text -ne $null -or $group.txt -ne $null ) {
$find_Group="*Staff-$Group.text -$dept.text*"
}
Staff-<Group>-<Unit>
get-ADobject -filter { name -like $Women_Location } | add-adgroupmember -members $sAMAcctName
}
###########################################################################dropdown boxes (office and function)

<#
$title=new-object System.Windows.Forms.Combobox
$title.Location = New-Object System.Drawing.Size($COL1,170)#
$title.Size = New-Object Drawing.Size(230,200)
#sets the height of the drop down
$title.DropDownHeight=200
#check to see if the Job titles file exists:
try { get-item $PSScriptRoot\Job_Titles.txt | Out-Null }
catch {
write-host "Unable to find file Job_Titles.txt"
Exit
}
#populate the values here from a file
$count=0
ForEach($i in (Get-Content $PSScriptRoot\Job_Titles.txt ))
{ $title.items.add($i) | Out-Null
$count++ | Out-Null
}
$title.SelectedIndex=($count/2)
$winform.controls.add($title)
#Function
$FN_Label = add-label -name "OU:" -XLoc $COL1 -YLoc 200
$winform.controls.add($FN_Label)
#description and jobtitle (get a list from AD?
#>
<#
$Department=new-object System.Windows.Forms.Combobox
$Department.Location = New-Object System.Drawing.Size($COL1,270)#
$Department.size = New-Object Drawing.Size(230,200)
#check to see if the Job titles file exists:
try { get-item $PSScriptRoot\Departments.txt | Out-Null }
catch {
write-host "Unable to find file Departments.txt"
Exit
}
#populate the values here from AD
$count=0
ForEach($i in (Get-Content $PSScriptRoot\Departments.txt))
{ $Department.items.add($i) | Out-Null
$count++ | Out-Null
}
#sets the height of the drop down
$Department.DropDownHeight=200
# select a value half way down the list
# $Department.SelectedIndex=($count/2)
$winform.controls.add($Department)
#>
<#
##################################combo boxes (multiselection) Security and email groups
#security groups
# column 2
$FN_Label = add-label -name "Security Groups:" -XLoc $COL2 -YLoc 100 -width 120
$winform.controls.add($FN_Label)
$Security_groups=new-object System.Windows.Forms.Listbox
$Security_groups.Location = New-Object System.Drawing.Size($col2,120)#
$Security_groups.Size = New-Object Drawing.Size(230,200)
#populate the values here from AD
ForEach($i in ($Security_Group_List))
{ $Security_groups.items.add($i.name) | Out-Null }
$Security_groups.SelectionMode = 'MultiExtended'
#$Security_groups.SelectedValue = ''
$winform.controls.add($Security_groups)
# column 3
#Distro groups
$FN_Label = add-label -name "Distro Groups:" -XLoc $COL3 -YLoc 100 -width 100
$winform.controls.add($FN_Label)
$Distro_groups=new-object System.Windows.Forms.Listbox
$Distro_groups.Location = New-Object System.Drawing.Size($COL3,120)#
$Distro_groups.size = New-Object Drawing.Size(230,200)
#populate the values here from AD
ForEach($i in ($Distribution_Group_List))
{ $Distro_groups.items.add($i.name) | Out-Null }
$Distro_groups.SelectionMode = 'MultiExtended'
$winform.controls.add($Distro_groups)
#>

