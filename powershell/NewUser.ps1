##### A script to create a new user.###

#
# Creates a new user account in AD, includiing security/distribution groups and puts them in the right OU.
# Creates a new user in Office 365 using the old CreateMailbox script.
# Sends an email to the user and their manager with some helpful links
# sends an email to App Support if there is an XRM or Rito account needed
#    (which could be done with powershell later on, when we move to Dynamics Online)
###

<#Updates for version 14
[*] Get a list of distribution groups from Office 365
get-distributiongroup ~ WHERE {$_.HiddenFromAddressListsEnabled -eq $false } | | select name

[*] If the box for XRM/Rito user is ticked, make sure something is typed in the box.
[*] Update New User script to populate the Province field with a combination of the Deparment and Office fields
[*] Quit the script if creating an 0365 account takes too long
[] "Copy account" option
[*]mkdir for home dir if needed
[] Additonal mailboxes (search, via email address)
#>

# version history
# v13:
<#
Update new user form to pull email groups from 0365
[*] Add user to # XRM / #Rito users list if ticked.

#>
# v12 - Bug fixes ( using > instead of -gt)
# v11 - Bug fixes
# v10 - Bug fixes
# v9 - um...
# v8 - um...
# v7 - Change some stuff to a function
# v6 - add progress bars and updates and handle errors niceish
# v5 - Generate and send an email
# modify other elememts (GUI or repeated code) into functions
# V4 - create an O365 account
# v3 - Create an AD user with the script, modified the "label" to be a function.
# v2 - GUI variables echo out to test them
# v1 - created the GUI

# Declare a global variable

# declare some functions

# this function adds the text labels into the GUI
function add-label {
     # these are the function's arguments
     param( [string]$name, [int]$XLoc, [int]$YLoc, [int]$width)

     $Label = New-Object System.Windows.Forms.Label
     # put it somewhere
     $Label.Location = New-Object System.Drawing.Size($XLoc,$YLoc)
     # this is the text of the label
     $Label.Text = $name

     # sets the width of the text label
     if ($width -eq "") { $width = 80 }
     $Label.Size = New-Object Drawing.Size($width,20)

     #return the object to use it
     return $label
 }

 # Add an alert box
 function add-AlertDialog {
    # these are the function's arguments
    param( [string]$title, [string]$message)
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

 # update the progress bar
 function update-progress {
   param( [string]$title,[int]$percent,[System.Windows.Forms.ProgressBar]$PB,[System.Windows.Forms.Label]$ProgLabel,[System.Windows.Forms.Form]$form)
   $PB.Value = $percent
   $ProgLabel.Text = $title
   $form.Refresh()
   Start-Sleep -Milliseconds 150
  }

 #
 # Generate a progress bar for loading the AD stuff.
 #

 function add-progressbar {

 param([string]$title)
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |  out-null
 Add-Type -assembly System.Windows.Forms
 $Form = new-object system.Windows.Forms.Form

 ## -- Create The Progress-Bar

 $Form.TopMost = $True
 $Form.Text = $title
 $Form.Height = 100
 $Form.Width = 500
 $Form.BackColor = "White"

 $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
 $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

 ## -- Create The text Label
 $ProgLabel = New-Object System.Windows.Forms.Label
 $ProgLabel.Text = "Starting. Please wait ... "
 $ProgLabel.Left = 5
 $ProgLabel.Top = 10
 $ProgLabel.Width = 500 - 20
 $ProgLabel.Height = 15
 $ProgLabel.Font = "Tahoma"
 ## -- Add the label to the Form
 $Form.Controls.Add($ProgLabel)

  #set up the progress bar
 $PB=New-Object System.Windows.Forms.ProgressBar
 $PB.Name = "PowerShellProgressBar"
 $PB.Value = 0
 $PB.Style="Continuous"

 $System_Drawing_Size = New-Object System.Drawing.Size
 $System_Drawing_Size.Width = 500 - 40
 $System_Drawing_Size.Height = 20

 $PB.Size = $System_Drawing_Size
 $PB.Left = 5
 $PB.Top = 40
 $Form.Controls.Add($PB)

 ## -- Show the Progress-Bar and Start The PowerShell Script
 $Form.Show() | Out-Null
 $Form.Focus() | Out-NUll
 $ProgLabel.Text = "Starting. Please wait ... "
 $Form.Refresh()

  return $PB,$ProgLabel,$form
 }

 #
 # set up some variables
 #

 # clear any errors up so the variable is fresh
 $error.clear()

 #Get-Variable |  where {$sysvars -notcontains $_.Name} | foreach {Remove-Variable $_}

 # sets the domain controller where the accounts are first made/manipulated.
 $DOMAIN_CONTROLLER="xxx"

 ## set some login information
$AD_Username = 'xxx'
try { $encrypted = Get-Content -Path 'C:\ScheduledTasks\ADUser_Encrypted_password.txt'}
catch { add-AlertDialog -title "Loading Password file" -message "Unable to load password file at C:\ScheduledTasks\ADUser_Encrypted_password.txt"
    Exit }
$key = (1..16)
$AD_Password = $encrypted | ConvertTo-SecureString -Key $key
$AD_Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AD_Username, $AD_Password

$Username = 'xxx'
try { $encrypted = Get-Content -Path 'C:\ScheduledTasks\Office365Admin_Encrypted_password.txt'}
catch { add-AlertDialog -title "Loading Password file" -message "Unable to load password file at C:\ScheduledTasks\Office365Admin_Encrypted_password.txt"
    Exit }

$key = (1..16)
$Password = $encrypted | ConvertTo-SecureString -Key $key
$O365_Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password

#Import MS Online Cmdlets
try{
    Import-module MSOnline
    Write-host -ForegroundColor Green "Importing MsOnline cmdlets"
}
Catch{
    Write-Host -ForegroundColor Red "Could not load MsOnline module. Ensure it is installed."
    add-AlertDialog -title "Importing MsOnline cmdlets" -message "Could not load MsOnline module. Ensure it is installed"
    $P_bar[2].Close()
    Exit
}

#Connect to Microsoft Online
try{
    connect-msolservice -credential $O365_Credentials -ErrorAction Stop
    Write-Host -ForegroundColor Green "Connecting to MS Online"
}
Catch{
    Write-Host -ForegroundColor Red "Could not connect to the service. Ensure the credentials are correct then try again."
    add-AlertDialog -title "Connecting to MS Online" -message "Could not connect to the service. Ensure the credentials are correct then try again."
    $P_bar[2].Close()
    Exit
}

try {
    $CloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365_Credentials -Authentication Basic –AllowRedirection
    Import-PSSession $CloudSession -AllowClobber | out-null
    }
catch {
    # add-alert
    write-host "Unable to connect to Office 365, please check the user name and password"
    add-AlertDialog -title "Connect to Office 365" -message "Unable to connect to Office 365, please check the user name and password"
    $P_bar[2].Close()
    exit
}

$P_Bar=add-progressbar -title "Loading Active Directory information"

 Start-Sleep -Seconds 1

 # stuff to load goes here
 #
 #

#
# Load up some information from Active Directory/Office 365
#

# the list of Offices (OU list, really)
 try{ $Office_List=get-adorganizationalunit -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel  -filter { name -notlike "*z*"} -Credential $AD_credentials | select name
      update-progress -title "Loading Active Directory information" -percent 20 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
 catch { write-host "unable to get list of offices" }

# list of Folder security groups
 try{ $Security_Group_List=get-adgroup -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel  -filter { name -notlike "*z*"} -Credential $AD_credentials | select name
      update-progress -title "Loading Active Directory information" -percent 40 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
 catch {  write-host "unable to get list of security groups" }

 # Distribution lists from AD

 #try{ $Distribution_Group_List=get-adgroup -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel  -filter { name -notlike "*z*"} -Credential $AD_credentials | select name
 #     update-progress -title "Loading Active Directory information" -percent 60 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
 #   }
 #catch { write-host "unable to get list of distro groups" }

# Distribution lists from Office 365
try{ $Distribution_Group_List=get-distributiongroup | WHERE {$_.HiddenFromAddressListsEnabled -eq $false }  | select name
     update-progress -title "Loading Active Directory information" -percent 60 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
   }
catch { write-host "unable to get list of distro groups" }

# Sharepoint security groups
try{ $Sharepoint_Group_List=get-adgroup -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel  -filter { name -notlike "*z*"} -Credential $AD_credentials | select name
     update-progress -title "Loading Active Directory information" -percent 80 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
 catch { write-host "unable to get list of Sharepoint groups" }

# Printer security groups
try{ $Printer_Group_List=get-adgroup -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel  -filter { name -notlike "*z*"} -Credential $AD_credentials| select name
     update-progress -title "Loading Active Directory information" -percent 100 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
 catch { write-host "unable to get list of Printer groups" }

 <#
 $Counter = 0
 ForEach ($Item In $Office_List) {
  ## -- Calculate The Percentage Completed
  $Counter++
        [Int]$Percentage = ($Counter/$Office_List.Count)*100

        update-progress -title "Loading Active Directory information" -percent $Percentage -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

}#>

$P_bar[2].Close()

#
# Draw the form, including the loaded information from AD.
#

## draw the form itself:
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |  out-null
 $WinForm = new-object Windows.Forms.Form
 $WinForm.TopMost = $True
 # set the name of the window:
 $WinForm.text = "New User Creation"
 #set the size of the window: Y by X
 $WinForm.Size = new-object Drawing.Size(750,600)
 #$WinForm.AutoSize = $True
 $WinForm.AutoSizeMode = "GrowAndShrink"
 $Font = New-Object System.Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Regular)
 # Font styles are: Regular, Bold, Italic, Underline, Strikeout
 $WinForm.Font = $Font

 # set up the gui controls"
 #$ListBox = new-object Windows.Forms.ListBox

 # text spacing is 30 (start at 15), text boxes 30 (start at 10)

 # set the width of the main 3 columns
 $COL1=5
 $COL2=250
 $col3=500

 $FN_Label = add-label -name "First name:" -XLoc $COL1 -YLoc 15
 $winform.controls.add($FN_Label)

 $First_Name=new-object System.Windows.Forms.Textbox
 $First_Name.Location = New-Object System.Drawing.Size(90,10)
 $First_Name.Size = New-Object System.Drawing.Size(100,10)
 $winform.controls.add($first_name)

 # Last name

 $FN_Label = add-label -name "Last name:" -XLoc 200 -YLoc 15
 $winform.controls.add($FN_Label)

 $Last_Name=new-object System.Windows.Forms.Textbox
 $Last_Name.Location = New-Object System.Drawing.Size(280,10)
 $Last_Name.Size = New-Object System.Drawing.Size(100,10)
 $winform.controls.add($Last_name)

 ###########################################################################Copy account option

 # Copy account

 #functions
 $DupeAcct_Ticked={
  write-host "running function"
# only check the text of the box if the box is ticked.
if ($script:DupeAcct_Check.CheckState -eq "checked") {
   $script:DupeAcct.add_lostfocus($DupeAcct_select)
   #disable the other controls
   $script:Office.enabled=$false
   $script:print_groups.enabled=$true
   $script:ShareP_groups.enabled=$true
   $script:Distro_groups.enabled=$true
   $script:Security_groups.enabled=$true
   $script:title.enabled=$true
   $script:function_text.enabled=$true
   }
if ($script:DupeAcct_Check.CheckState -eq "unchecked") {
   #enable the other controls
   $script:Office.enabled=$true
   $script:print_groups.enabled=$true
   $script:ShareP_groups.enabled=$true
   $script:Distro_groups.enabled=$true
   $script:Security_groups.enabled=$true
   $script:title.enabled=$true
   $script:function_text.enabled=$true
   }
}

 # Dupe  user tickbox (the text is self contained)
 $DupeAcct_Check = New-Object System.Windows.Forms.Checkbox
 $DupeAcct_Check.Location = New-Object System.Drawing.Size($COL1,50)
 $DupeAcct_Check.Size = New-Object System.Drawing.Size(140,20) # to show the text
 $DupeAcct_Check.Text = "Copy AD Acct:"
 $DupeAcct_Check.TabIndex = 3

 $DupeAcct_Check.Add_CheckStateChanged($DupeAcct_Ticked)

 $winform.Controls.Add($DupeAcct_Check)

 $DupeAcct=new-object System.Windows.Forms.Textbox
 $DupeAcct.Location = New-Object System.Drawing.Size(150,45)
 $DupeAcct.Size = New-Object System.Drawing.Size(100,10)

  $DupeAcct_Select={
  $DupeAcct_text=""
  write-host "updating account copy text box"
  $script:DupeAcct_find=$DupeAcct.text
  #debugging
  #$DupeAcct_find=$DupeAcct.text
  write-host "updating Duplicating Account to: $DupeAcct_find"
  # search for the copy's name, error alert if not found.
  $DupeAcct_text = get-aduser -properties * -filter { name -like $DupeAcct_find } -SearchBase "xxx" -Credential $AD_credentials
  write-host $DupeAcct_text.name

  if ($DupeAcct_text) { write-host "found account" $DupeAcct_text.mail
      }
  else {  write-host "Account doesn't exist!"
      add-AlertDialog -title "Account " -message "The account you entered for was not found. Check your spelling!"
    }
  }

$DupeAcct.add_lostfocus($DupeAcct_select)
$winform.controls.add($DupeAcct)





 ###########################################################################dropdown boxes (office and function)

 #$Office

 $FN_Label = add-label -name "Office:" -XLoc $COL1 -YLoc 100
 $winform.controls.add($FN_Label)

 $Office=new-object System.Windows.Forms.Combobox
 $Office.Location = New-Object System.Drawing.Size($COL1,120)#
 $Office.Size = New-Object Drawing.Size(230,20)

 #populate the values here from AD
 $count=0
 ForEach($i in $Office_List)
  { $Office.items.add($i.name) | Out-Null
    $count++  | Out-Null
  }

 #sets the height of the drop down
 $Office.DropDownHeight=200
 $Office.SelectedIndex=($count/2)

##########
 $office_Select={
  write-host "updating combo box"
  $script:SearchOU=$Office.text
  write-host "xxx"
  $FUNC_ARRAY=get-adorganizationalunit -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel -filter *  -Credential $AD_credentials | select name
  $Function.Items.Clear()
  ForEach($FUNC in $FUNC_ARRAY) { $Function.items.add($FUNC.name) | Out-Null }
  $function.SelectedIndex=(1)
  }

 $office.add_SelectedIndexChanged($Office_select)

 $winform.controls.add($Office)


 # job title

 $FN_Label = add-label -name "Job title:"  -XLoc $COL1 -YLoc 150
 $winform.controls.add($FN_Label)

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

 $FN_Label = add-label -name "OU:"  -XLoc $COL1 -YLoc 200
 $winform.controls.add($FN_Label)

 $Function=new-object System.Windows.Forms.Combobox
 $Function.Location = New-Object System.Drawing.Size($COL1,220)#
 $Function.size = New-Object Drawing.Size(230,20)

 #Values from the array:
 $count=0

 $SearchOU=$Office.text
 $FUNC_ARRAY=get-adorganizationalunit -server $DOMAIN_CONTROLLER -searchbase "xxx" -searchscope OneLevel -filter * -Credential $AD_credentials | select name
 ForEach($FUNC in $FUNC_ARRAY) { $Function.items.add($FUNC.name) | Out-Null
   $count++ | Out-Null
   }

 $function.SelectedIndex=0 # default to "Red Cross"
 $winform.controls.add($Function)

 #description and jobtitle (get a list from AD?

 $FN_Label = add-label -name "Department:"  -XLoc $COL1 -YLoc 250 -width 100
 $winform.controls.add($FN_Label)

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

 # Tick boxes
 #

 # XRM user tickbox (the text is self contained)
 $XRM = New-Object System.Windows.Forms.Checkbox
 $XRM.Location = New-Object System.Drawing.Size($COL1,300)
 $XRM.Size = New-Object System.Drawing.Size(100,20) # to show the text
 $XRM.Text = "XRM User"
 $XRM.TabIndex = 3
 $winform.Controls.Add($XRM)

  # Rito User tickbox (the text is self contained)

 $RITO = New-Object System.Windows.Forms.Checkbox
 $RITO.Location = New-Object System.Drawing.Size($COL1,330)
 $RITO.Size = New-Object System.Drawing.Size(100,20)
 $RITO.Text = "RITO user"
 $RITO.TabIndex = 4
 $winform.Controls.Add($RITO)

 # Rito

 $FN_Label = add-label -name "XRM/Rito Acct to cpy:"  -XLoc $COL1 -YLoc 360 -width 120
 $winform.controls.add($FN_Label)

 $Copy_account=new-object System.Windows.Forms.Textbox
 $Copy_account.Location = New-Object System.Drawing.Size(130,360)
 $Copy_account.Size = New-Object System.Drawing.Size(100,10)

 $winform.controls.add($Copy_account)

 # Manager

 $FN_Label = add-label -name "Manager:"  -XLoc $COL1 -YLoc 390
 $winform.controls.add($FN_Label)

 $manager=new-object System.Windows.Forms.Textbox
 $manager.Location = New-Object System.Drawing.Size(90,390)
 $manager.Size = New-Object System.Drawing.Size(100,10)

  $Manager_Select={
  $manager_text=""
  write-host "updating Manager box"
  $script:manager_find=$manager.text
  #debugging
  #$manager_find=$manager.text
  write-host "updating Manager to: $manager_find"
  # search for the managers name, error alert if not found.
  $manager_text = get-aduser -properties * -filter { name -like $manager_find } -SearchBase "xxx" -Credential $AD_credentials
  write-host $manager_text.name

  if ($manager_text) { write-host "found manager" $manager_text.mail
      }
  else {  write-host "Manager doesn't exist!"
      add-AlertDialog -title "Manager" -message "The manager you entered for was not found. Check your spelling!"
    }

  }

$manager.add_lostfocus($Manager_select)
$winform.controls.add($manager)

$OKButton=New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point($col1,420)
$OKButton.Size = New-Object System.Drawing.Size(100,23)
$OKButton.Text = 'Create User'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$winform.AcceptButton=$OKButton
$winform.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point($col1,450)
$CancelButton.Size = New-Object System.Drawing.Size(100,23)
$CancelButton.Text = 'Quit'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$winform.CancelButton = $CancelButton
$winform.Controls.Add($CancelButton)

 ##################################combo boxes (multiselection) Security and email groups

 #security groups

 # column 2

 $FN_Label = add-label -name "Security Groups:"  -XLoc $COL2 -YLoc 100  -width 120
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

 $FN_Label = add-label -name "Distro Groups:"  -XLoc $COL3 -YLoc 100  -width 100
 $winform.controls.add($FN_Label)

 $Distro_groups=new-object System.Windows.Forms.Listbox
 $Distro_groups.Location = New-Object System.Drawing.Size($COL3,120)#
 $Distro_groups.size = New-Object Drawing.Size(230,200)

  #populate the values here from AD
 ForEach($i in ($Distribution_Group_List))
 { $Distro_groups.items.add($i.name) | Out-Null }

 $Distro_groups.SelectionMode = 'MultiExtended'
 $winform.controls.add($Distro_groups)


##################################combo boxes (multiselection) printing and sharepoint

# combobox for printing security groups, 5 lines high
#Add a security group for offices that will never have follow-me printing?
# skip the direct printing SG

 # column 3

 $FN_Label = add-label -name "Printing Groups:"  -XLoc $COL2 -YLoc 330  -width 120
 $winform.controls.add($FN_Label)

 $print_groups=new-object System.Windows.Forms.Listbox
 $print_groups.Location = New-Object System.Drawing.Size($col2,350)#
 $print_groups.Size = New-Object Drawing.Size(230,200)

 #populate the values here from AD
 ForEach($i in ($Printer_Group_List))
  { $print_groups.items.add($i.name) | Out-Null }

 $print_groups.SelectionMode = 'MultiExtended'
  $winform.controls.add($print_groups)

 # column 3
 # also sharepoint groups, try to add a default group based on the department?

 $FN_Label = add-label -name "Sharepoint Groups:"  -XLoc $COL3 -YLoc 330  -width 160
 $winform.controls.add($FN_Label)

 $ShareP_groups=new-object System.Windows.Forms.Listbox
 $ShareP_groups.Location = New-Object System.Drawing.Size($col3,350)#
 $ShareP_groups.size = New-Object Drawing.Size(230,200)

  #populate the values here from AD
 ForEach($i in ($SharePoint_Group_List))
 { $ShareP_groups.items.add($i.name) | Out-Null }

 $ShareP_groups.SelectionMode = 'MultiExtended'
 $winform.controls.add($ShareP_groups)


#display the form:
$WinForm.Add_Shown($WinForm.Activate())
# $WinForm.showdialog() | out-null

$result = $WinForm.ShowDialog()

if ($result -eq "Cancel") {
    Write-Host -ForegroundColor Green "Tidying Up"
    Remove-PSSession $CloudSession
    Remove-Module MSOnline
    Remove-Module ActiveDirectory
    Exit
    }
if ($result -eq "Ok")
    {
#
# Some error checking on the form
#

# users need a manager, and first and last names

if ( $First_Name.text -eq "" ) { add-AlertDialog -title "First name cannot be blank" -message "You can't create a new user without a first name."
    exit
    }

if ( $Last_Name.text -eq "" ) { add-AlertDialog -title "Last name cannot be blank" -message "You can't create a new user without a last name."
    exit
    }

if ( $Copy_account.text -eq "" -and ($XRM.CheckState -eq "checked" -or $RITO.CheckState -eq "checked"))
    { add-AlertDialog -title "XRM/RITO account name cannot be blank" -message "You can't copy a new XRM or RITO account without a name."
    exit
    }

if ($DupeAcct_Check.CheckState -eq "checked" -and $DupeAcct_text -ne "" ) {
    # get the following information from the duplicated account

    $Department.text=$DupeAcct_text.department
    $title.Text = $DupeAcct_text.title
    foreach ( $ADgroup in $DupeAcct_text.memberof ) {
         add-adgroupmember -members "$first_name_text.$last_name_text" -identity $ADgroup -server $DOMAIN_CONTROLLER -credential $AD_Credentials
         }
    # OU to put them in
    $DupeAcct_text.CanonicalName
    $OUARRAY=$DupeAcct_text.CanonicalName.("/")
    $function_text=$OUARRAY[($OUARRAY.count -1)]
    # also the same email groups:
    $DistributionGroups=Get-DistributionGroup | where { (Get-DistributionGroupMember $_.DistinguishedName | foreach {$_.PrimarySmtpAddress}) -contains "$Username"} | select DistinguishedName
    $Distro_groups.SelectedItems=$DistributionGroups


       }
    }


#
# Active Directory : Create the account
#

# Show another progress log dialog
#

$P_Bar=add-progressbar -title "Setting up account"

#generate a password
#there are better ways to generate a pasword than this

update-progress -title "Generating password" -percent 5 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

$P1=get-random -Maximum 12 -Minimum 0
$P2=get-random -Maximum 12 -Minimum 0
$P3=get-random -Maximum 9 -Minimum 0
$P4=get-random -Maximum 9 -Minimum 0

$PASSWORD_ARRAY=@(
    "Summner", "Spring","Winter","Autumn",
    "Green","Good","Integrity","Improve","Unity",
    "Humanity","Society","Blue","Orange"
    )

# generate a random paassword
#write-host $p1 $p2 $p3 $p4
$Bit_1=$PASSWORD_ARRAY[$P1]
$Bit_2=$PASSWORD_ARRAY[$P2]
$PASSWORD="$Bit_1$Bit_2$p3$p4"


#$PASSWORD_ARRAY[$P2]$P3$P4"
#write-host $password

# create the account
# stuff you need:
# h drive, first name, last name, displayname, UPN, OU, password, change password at login, physicaloffice, manager, address, title/description, department,
# company, default group memberships

# do this because the strings don't like dot syntax much.
$first_name_text=$first_name.text
$last_name_text=$last_name.text
$function_text=$Function.Text
$Office_text=$Office.text
$CONVERT_PASSWORD=ConvertTo-SecureString -string $password -asplaintext -force

# do this because I'm not sure how to get the varables back out of the function for the manager search
$manager_find=$manager.text
$manager_text = get-aduser -properties * -filter { name -like $manager_find } -SearchBase "OU=Fake Users,DC=corp,DC=company,DC=org" -Credential $AD_credentials

$NewUSerArguments = @{
    GivenName = $first_name_text
    Surname = $last_name_text

    DisplayName = "$first_name_text $last_name_text"
    Name = "$first_name_text $last_name_text"
    SamAccountName = "$first_name_text.$last_name_text"

    UserPrincipalName = "$first_name_text.$last_name_text@company.org"
    path = "OU=$function_text,OU=$Office_text,OU=NZRC Users,DC=corp,DC=company,DC=org"

    EmailAddress = "$first_name_text.$last_name_text@company.org"
    HomeDirectory = "\\cloudfileservername\home\$first_name_text.$last_name_text"
    HomeDrive = "H:"
    # AD object
    #Manager = $manager_text

    Office = $Office.text
    Company = "Company"
    Department = $Department.text
    Description = $title.Text
    Title = $title.Text

    AccountPassword = $CONVERT_PASSWORD
    enabled = $true
    ChangePasswordAtLogon = $true

    server = $DOMAIN_CONTROLLER
    credential = $AD_Credentials
    }

    # use this block for debugging the AD user creation.
    write-host first: $first_name_text
    write-host last: $last_name_text

    write-host display name: "$first_name_text $last_name_text"
    write-host name : "$first_name_text $last_name_text"
    write-host SAM : "$first_name_text.$last_name_text"

    write-host UPN: "$first_name_text.$last_name_text@company.org"
    write-host path:  "OU=$function_text,OU=$Office_text,OU=Users,DC=corp,DC=company,DC=org"

    write-host email: "$first_name_text.$last_name_text@comnpany.org"
    write-host home dir: "\\fileservername\home\$first_name_text.$last_name_text"
    write-host home drive: "H:"
    write-host manager: $manager_text
    write-host office: $Office.text
    write-host company: "Company"
    write-host dept: $Department.text
    write-host title:  $title.Text
    write-host desc: $title.Text
    write-host state: "$Office.text $Department.text"

    write-host pwd: $CONVERT_PASSWORD
    write-host dc: $DOMAIN_CONTROLLER
    write-host dcpwd: $AD_Credentials

write-host "Creating user account: $first_name_text $last_name_text"
update-progress -title "Creating user account: $first_name_text $last_name_text" -percent 10 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

new-aduser @NewUSerArguments

# keep this for debugging
# new-aduser -UserPrincipalName "$first_name_text.$last_name_text@company.org" -Manager $manager_text -GivenName $first_name_text -Surname $last_name_text -DisplayName "$first_name_text $last_name_text" -Name "$first_name_text $last_name_text" -SamAccountName "$first_name_text.$last_name_text"  -path "OU=$function_text,OU=$Office_text,OU=Users,DC=corp,DC=company,DC=org" -EmailAddress "$first_name_text.$last_name_text@company.org"   -HomeDirectory "\\fileservername\home\$first_name_text.$last_name_text" -HomeDrive "H:" -Office $Office.text -Company "Company" -Department $Department.text -Description $title.Text -Title $title.Text -AccountPassword $CONVERT_PASSWORD  -enabled $true -ChangePasswordAtLogon $true -server $DOMAIN_CONTROLLER -credential $AD_Credentials

# of the user exists, this will spit up an error message.
if ( $error[0] ) {
   if ( $error[0] -like "*Unknown error (0x21c8)*" ) { add-AlertDialog -title "Account Creation" -message "A user account with this name already exists" }
   else {add-AlertDialog -title "Account Creation" -message $error[0].ToString() }
   $P_bar[2].Close()
   #exit
   }

<# should't need this as creating it on the pdc
Try{
    Write-Host -ForegroundColor Green "Syncing Active Directory"
    invoke-command -credential $credentials -computername $DOMAIN_CONTROLLER -scriptblock { & cmd /c C:\Windows\System32\repadmin.exe /syncall /AdePq }
}
Catch{
    Write-Host -ForegroundColor Red "Unable to sync Active Directory."
    Exit
}#>

# add to security groups, security groups, sharepoint groups and other groups
# find out if we can add the users directly

# these should be a function really....

update-progress -title "Adding to AD groups" -percent 20  -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

 # update the security groups
 foreach ( $group in $Security_groups.Selecteditems ) {
    write-host "Adding to $group"
    add-adgroupmember -identity $group -members "$first_name_text.$last_name_text" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
    }

 #foreach ( $group in $Distro_groups.SelectedItems ) {
 #   write-host $group
 #   add-adgroupmember -identity $group -members "$first_name_text.$last_name_text" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
 #   }

# update the sharepoint groups
foreach ( $group in $ShareP_groups.SelectedItems ) {
    write-host "Adding to $group"
    add-adgroupmember -identity $group -members "$first_name_text.$last_name_text" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
    }

# update the printer groups
foreach ( $group in $print_groups.SelectedItems ) {
    write-host "Adding to $group"
    add-adgroupmember -identity $group -members "$first_name_text.$last_name_text" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
    }

# add to the default groups
# these are just examples
#
write-host "Adding to: SG_Remote_Desktop, All Users, SG-folder-companyDrive"
add-adgroupmember -members "$first_name_text.$last_name_text" -identity "SG_Remote_Desktop" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
add-adgroupmember -members "$first_name_text.$last_name_text" -identity "All Users" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
add-adgroupmember -members "$first_name_text.$last_name_text" -identity "SG_Folder-CompanyDrive" -server $DOMAIN_CONTROLLER -credential $AD_Credentials

# if the Xrm or Rito boxes are ticked, add to the appropriate list
# these will need to be updated if the groups go to Office 365

# has been moved until after the 365 user account is created
#if ( $XRM.CheckState -eq "checked" ) {
#add-groupmember -members "$first_name_text.$last_name_text" -identity "XRMUsers" -server $DOMAIN_CONTROLLER -credential $AD_Credentials
#}

# Per location switch for

   switch ( $Office )
    {

       "Christchurch Recovery" { write-host "Not adding to group" }
        default { write-host "Adding toi group"
          add-adgroupmember -members "$first_name_text.$last_name_text" -identity "GroupName" -server $DOMAIN_CONTROLLER -Credential $AD_credentials
          }
    }

# sync all of the DC's before trying the Office 365 stuff.

#
# Office 365
#


$userInput="$first_name_text.$last_name_text@company.org"

<#
Try{
    Write-Host -ForegroundColor Green "Importing Active Directory cmdlets"
    Import-Module ActiveDirectory
}
Catch{
    Write-Host -ForegroundColor Red "Unable to Import Active Directory Module. Ensure it is installed."
    Exit
}#>

#Write-Host -ForegroundColor Green "Getting AD User" $userInput

<#
if ( Get-ADUser -Filter { userprincipalname -like $userInput } ) {
$ADUser = Get-ADUser -Filter { userprincipalname -like $userInput }
Write-Host -ForegroundColor Green "AD user account:" $ADUser
}
else {
    Write-Host -ForegroundColor Red "Unable to find user in AD. Ensure that they exist and that their UPN is correct"
    Write-Host -ForegroundColor Red "If you have created the account on any DC other than the cloud DC, it may take up to 10 mins to sync"
    Exit
}
#>

update-progress -title "Creating Office 365 Account" -percent 30 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

# if the user is a delegate or an instructor, assign an E2, otherwise an e3
if ( $function -eq "Delegates" -or $function -eq "Instructor" ) { $userLicenseValue = 'E2' }
Else { $userLicenseValue = 'E3' }

###
#
#write out the variables and set the SMTP etc for the user
#
###
$ADUser = Get-ADUser -Filter { userprincipalname -like $userInput }

Write-Host -ForegroundColor Green 'sam account name'+$ADUser.SamAccountName
$MailNickName = $ADUser.Name
$primarySMTP = 'SMTP:'+$ADUser.SamAccountName+'@company.org'
$emailAddress = $ADUser.SamAccountName+'@company.org'
$proxyAddress = 'smtp:'+$ADUser.SamAccountName+'@company.mail.onmicrosoft.com'
$SIPAddress = 'sip:'+$ADUser.SamAccountName+'@company.org'
Write-Host -ForegroundColor Green "Updating proxy address with" $proxyAddress " & " $primarySMTP " & " $SIPaddress

Set-ADUser $ADUser -Add @{proxyAddresses=$primarySMTP,$proxyAddress,$SIPaddress} -credential $AD_Credentials -server $DOMAIN_CONTROLLER
Set-ADUser $ADUser -Add @{displayName=$ADUser.Name} -credential $AD_Credentials -server $DOMAIN_CONTROLLER
Set-ADUser $ADUser -Add @{mailNickName=$ADUser.Name} -credential $AD_Credentials -server $DOMAIN_CONTROLLER

# add the home folder if is doesn't exist yet

if(get-item -literalpath "\\cloudfileservername\home\$first_name_text.$last_name_text") { write-host "Home Directory exists" }
else {
    new-item -ItemType directory -path "\\cloudfileservername\home\$first_name_text.$last_name_text"
    #$perm=$aduser,"FullControl", "ContainerInherit, objectinherit","none","allow"
    $acl= get-acl "\\cloudfileservername\home\$first_name_text.$last_name_text"
    $perm="$first_name_text.$last_name_text","FullControl", "ContainerInherit, objectinherit","none","allow"
    #$perm="corp\$first_name_text.$last_name_text","FullControl","allow"
    $rule=new-object -typename System.Security.AccessControl.FileSystemAccessRule -ArgumentList $perm
    $acl.setaccessrule($rule)
    $acl | set-acl -path "\\cloudfileservername\home\$first_name_text.$last_name_text"
    }


try{
 get-acl -path "\\cloudfileservername\home\$first_name_text.$last_name_text"
 }
Catch { add-AlertDialog -title "Home Directory" -message "Unable to create Home directory: \\cloudfileservername\home\$first_name_text.$last_name_text"
      }

try{
    Set-ADUser $ADUser -Add @{mail=$emailAddress } -credential $AD_Credentials -server $DOMAIN_CONTROLLER
    # -server $DOMAIN_CONTROLLER
}
Catch{
    Write-Host -ForegroundColor Red "Primary SMTP Address already set"
}

###
#
#Run DirSync to update O365 with the user's details
#
###

Write-Host -ForegroundColor Green "Running DirSync to copy updates to O365"
& "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -psconsolefile 'C:\Program Files\Microsoft Online Directory Sync\DirSyncConfigShell.psc1' -command "Start-OnlineCoexistenceSync"

#
#create the account in Office 365
#

#check to see if the user has been uploaded from AD
$CheckIfUserAccountExists = Get-MsolUser -UserPrincipalName $emailAddress -ErrorAction SilentlyContinue

$counter=30
do {
       $CheckIfUserAccountExists = Get-MsolUser -UserPrincipalName $emailAddress -ErrorAction SilentlyContinue
       #$a = get-date
       #Write-Host "Most recent check if account has been created yet at "$a.hour":"$a.minute":"$a.second
       update-progress -title "Creating account..." -percent $counter -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
       Write-Host "Creating account..."
       if ($counter -gt 40) {
         add-AlertDialog -title "Office 365 Account" -message "Unable to create Office 365 account"
         exit
         }
       $counter++
       Sleep 15
}
While ($CheckIfUserAccountExists -eq $Null)

###
#
# Set the Office 365 attributes
#
###

#set user location to NZ so license can be applied
Write-Host -ForegroundColor Green "Updating user location to NZ"
update-progress -title "Updating user location to NZ" -percent 50 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
Set-MsolUser -UserPrincipalName $emailAddress -UsageLocation NZ

wait-event -timeout 15

#add the license
if ($userLicenseValue -eq 'E3'){
    try{
        Set-MsolUserLicense -UserPrincipalName $userInput -AddLicenses company:ENTERPRISEPACK
        Write-Host -ForegroundColor Green "Added E3 License"
        update-progress -title "Added E3 License" -percent 55 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
    Catch{
        add-AlertDialog -title "Adding License" -message "Unable to add license. Ensure account user rights are set in O365 correctly"
        $P_bar[2].Close()
        Write-host -ForegroundColor Red "Unable to add license. Ensure account user rights are set in O365 correctly."
        Exit
    }
}
else {
    try{
        Set-MsolUserLicense -UserPrincipalName $userInput -AddLicenses company:STANDARDWOFFPACK
        Write-Host -ForegroundColor Green "Added E2 License"
        update-progress -title "Added E3 License" -percent 55 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    }
    Catch{
        add-AlertDialog -title "Adding License" -message "Unable to add license. Ensure account user rights are set in O365 correctly"
        $P_bar[2].Close()
        Write-host -ForegroundColor Red "Unable to add license. Ensure account user rights are set in O365 correctly."
        Exit
    }
}

#Give O365 a chance to create the mailbox...

$counter=60
$checkifmailboxexists = get-mailbox $emailAddress -erroraction silentlycontinue
do {
    $checkifmailboxexists = get-mailbox $emailAddress -erroraction silentlycontinue
    #$a = get-date
    #Write-Host "Checking if Mailbox has been created yet "$a.hour":"$a.minute":"$a.second
    update-progress -title "Creating mailbox..." -percent $counter -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
    if ($counter -gt 80) {
      add-AlertDialog -title "Office 365 Mailbox" -message "Unable to create mailbox"
      exit
      }
    $counter++
    Write-Host "Creating mailbox..."
    Sleep 15
    }
While ($checkifmailboxexists -eq $Null)

if (get-Mailbox -identity $userInput) {
    if ($userLicenseValue -eq 'E3'){
        try{
            Set-Mailbox -identity $userInput -IssueWarningQuota 49.50GB -ProhibitSendQuota 49.75GB -ProhibitSendReceiveQuota 50GB
            update-progress -title "Set mailbox size to 50 GB" -percent 90 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
            Write-Host -ForegroundColor Green "Set mailbox size to 50 GB"}
        Catch{
            add-AlertDialog -title "MailboxSize" -message "Unable to set mailbox size"
            $P_bar[2].Close()
            Write-host -ForegroundColor Red "Unable to set mailbox size"
            Exit }
        }
    else {
        try{
            Set-Mailbox -identity $userInput -IssueWarningQuota 4096MB -ProhibitSendQuota 4864MB -ProhibitSendReceiveQuota 5120MB
            update-progress -title "Set mailbox size to 5 GB" -percent 90 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]
            Write-Host -ForegroundColor Green "Set mailbox size to 5 GB" }
        Catch{
            add-AlertDialog -title "MailboxSize" -message "Unable to set mailbox size"
            $P_bar[2].Close()
            Write-host -ForegroundColor Red "Unable to set mailbox size"
            Exit }
        }

    Write-Host -ForegroundColor Green "Enabling Archive"
    Enable-Mailbox -identity $userInput -Archive
    }
else {
    #add-AlertDialog
    Write-Host -ForegroundColor Red "Unable to connect to O365 mailbox. Ensure that the user UPN is correct."
    Exit
    }

#Update membership to 365 groups

# update the office 365 groups

if ( $Distro_groups.SelectedItems -ne "" ) {
    write-host "Adding user to Office 365 groups:"
    foreach ( $group in $Distro_groups.SelectedItems ) {
       write-host $group
       add-distributiongroupmember -identity $group -members "$first_name_text.$last_name_text"
       }
    }

#add to the Rito and XRM user groups

if ( $XRM.CheckState -eq "checked" ) {
    write-host "Adding to XRM distribution group"
    add-distributiongroupmember -members "$first_name_text.$last_name_text" -identity "XRMUsers@company.org"
}

if ( $RITO.CheckState -eq "checked" ) {
   write-host "Adding to RITO distribution group"
   add-distributiongroupmember -members "$first_name_text.$last_name_text" -identity "RITOUsers@company.org"
}


##
# Email user and manager password and email details.
#
##

update-progress -title "Sending alert emails to mananger and user" -percent 100 -PB $P_Bar[0] -proglabel $P_Bar[1] -form $P_Bar[2]

# get user details for It.Services

$Email_Username = 'it.services@company.org'
$encrypted = Get-Content -Path 'c:\scripts\password.txt'
$key = (1..16)
$Email_Password = $encrypted | ConvertTo-SecureString -Key $key
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Email_Username, $Email_Password

$manager_email=$manager_text.mail
# ,$manager_email

$MailArgs = @{
   From       = "it.services@company.org"
   To         = $emailAddress
   Subject    = "Account Details for $First_Name_text $Last_Name_text"
   Body       = "Hi $First_Name_text,<br>
Your company computer username is: $First_Name_text.$Last_Name_text<br>
Your company temporary password is: $password<br>
Your company email address is $emailAddress<br>
<br>
Your printers, additional drives and any Outlook mailboxes will be attached to your account automatically.<br>
<br>
If you have any ICT issues, please call the ICT Service Desk<br>
<br>
Please have a look at the IT induction documents which will show you how to do things such as gaining remote access, setting up a mobile to access email, and accessing your email via the internet:<br>
<br>
<a href=https://company.sharepoint.com/ict/ICT%20Documents/Forms/AllItems.aspx?viewpath=%2Fict%2FICT%20Documents%2FForms%2FAllItems%2Easpx&id=%2Fict%2FICT%20Documents%2FHow%20To%20documents>
https://company.sharepoint.com/ict/ICT%20Documents/Forms/AllItems.aspx?viewpath=%2Fict%2FICT%20Documents%2FForms%2FAllItems%2Easpx&id=%2Fict%2FICT%20Documents%2FHow%20To%20documents
</a><br>
<br>
<br>
Thanks,<br>
The Company ICT Team<br>"
   SmtpServer = "smtp.office365.com"
   Port       = 587
   UseSsl     = $true
   Credential = $Credentials
}
Send-MailMessage @MailArgs -BodyAsHtml

$MailArgs = @{
   From       = "it.services@company.org"
   To         = $manager_email
   Subject    = "Account Details for $First_Name_text $Last_Name_text"
   Body       = "Hi $First_Name_text,<br>
Your Company computer username is: $First_Name_text.$Last_Name_text<br>
Your Company temporary password is: $password<br>
Your Company email address is $emailAddress<br>
<br>
Your printers, additional drives and any Outlook mailboxes will be attached to your account automatically.<br>
<br> <br>
<br>
Please have a look at the Red Cross IT induction documents which will show you how to do things such as gaining remote access, setting up a mobile to access Red Cross email, and accessing your email via the internet:<br>
<br>
<a href=https://newzealandredcross.sharepoint.com/ict/ICT%20Documents/Forms/AllItems.aspx?viewpath=%2Fict%2FICT%20Documents%2FForms%2FAllItems%2Easpx&id=%2Fict%2FICT%20Documents%2FHow%20To%20documents>
https://newzealandredcross.sharepoint.com/ict/ICT%20Documents/Forms/AllItems.aspx?viewpath=%2Fict%2FICT%20Documents%2FForms%2FAllItems%2Easpx&id=%2Fict%2FICT%20Documents%2FHow%20To%20documents
</a><br>
<br>
<br>
Thanks,<br>
The Company ICT Team<br>"
   SmtpServer = "smtp.office365.com"
   Port       = 587
   UseSsl     = $true
   Credential = $Credentials
}
Send-MailMessage @MailArgs -BodyAsHtml

# for testing
#$XRM.CheckState = "checked"
#$RITO.CheckState = "checked"

#If this is true then email the details to database admin
$SendDynamicsEmail=$false
if ( $XRM.CheckState -eq "checked" ) {
    $SendDynamicsEmail=$true
    $XRM_String="XRM"
    }
if ( $RITo.CheckState -eq "checked" ) {
   $SendDynamicsEmail=$true
   $XRM_String="Rito"
   }
if ( $RITo.CheckState -eq "checked" -and $XRM.CheckState -eq "checked" ) {
   $SendDynamicsEmail=$true
   $XRM_String="XRM and Rito"
   }
if ($Copy_account){
   $Copy_account_text="Copy this user account: $Copy_account."
   }

if ($SendDynamicsEmail -eq $true) {
    $MailArgs = @{
       From       = 'it.services@company.org'
       To         = 'databaseadmin@company.org'
       Subject    = "New user account setup for $UserInput"
       Body       = "Hi <br>Please set up an account for $UserInput in $XRM_String. $Copy_account_text<br>Thanks! <br>Company Service Desk"
       SmtpServer = 'smtp.office365.com'
       Port       = 587
       UseSsl     = $true
       Credential = $Credentials
    }
    Send-MailMessage @MailArgs -BodyAsHtml
}

$P_bar[2].Close()
#unload modules and tidy up.

Write-Host -ForegroundColor Green "Tidying Up"
Remove-PSSession $CloudSession
Remove-Module MSOnline
Remove-Module ActiveDirectory

#>
