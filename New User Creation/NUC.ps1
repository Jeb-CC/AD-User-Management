#=================================================================================================#
# TITLE: New User Creation (NUC, pronounced either as nuck or nook.)
# CREATOR: Ellis Svannish
# PURPOSE: Simplify the user creation process. Creates an Active Directory account, instructs technician 
#           to create mailbox in Exchange, and script awaits attachment of mailbox and automatically 
#           adjusts account mail values to meet domain requirements.
# DISCLAIMER: Should not be used as is, as it was created with a very specific process in mind. 
#             Should be modified to suit your suits, or taken as a form of inspiration.
# REQUIREMENTS: Admin access to Active Directory, ActiveDirectory PowerShell module, and modification 
#               of custom variables. See "UPDATE CUSTOM VARIABLES HERE" header below.
# FEATURES:
#       - Will check if a username is already taken before creation. (Refer to the :) or ): )
#       - Form options are responsive based on the OU you choose.
#       - Will automatically assign email domain, description, login script, and groups based on CSV.
#       - Compatible with middle name initials.
#       - Compatible with names that have spaces.
#       - Random password generator based on common dictionary words. Includes a button to
#		regenerate a new password.
#       - Creates network file share for the user using their M: drive.
#       - Automatically detects mail changes to account and corrects values to desired domain/name.
#=================================================================================================#

#This is to allow the Windows Message prompts to appear.
Add-Type -AssemblyName System.Windows.Forms

# F: Banner --------------------------------------------------------------------------------------#
function Banner {
	write-host
	write-host "Please see GUI to proceed."
	write-host "===================================================================="
	write-host
}
Banner

# UPDATE CUSTOM VARIABLES HERE
$scriptVersion = "NUC v3.4" #Version for display.
$dbPath = "D:\Scripts\UnitDatabase.csv" #File path of the Unit "Database".
$wordsPath = "D:\Scripts\Approved Words.txt" #File path of the dictionary word list for password generation.
$logoPath = "D:\Scripts\Logo.png" #File path of the company logo for display.
$onmicrosoftDomain = "@company.mail.onmicrosoft.com" #Onmicrosoft domain for emails.
$defaultDomain = "@company.com" #Default company domain for emails.

#==================#
#=== BEGIN FORM ===#
#==================#
# IMPORTS
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#=================================================================================================#
# FONTS
$font_header = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Bold)
$font_bold = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
#=================================================================================================#
# ICON
$icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
#=================================================================================================#
# FUNCTIONS

#This function will execute if the first name, last name, or initial is changed. It will update the display name and username.
function TextChanged {
	# Update Display Name
    # This function performs differently if initial is provided.
    if ($tb_initial.Text -eq "") {
        # Sets the display name variable to last name + , + first name.
        $tb_displayname.Text = $tb_lastname.Text + ", " + $tb_firstname.Text
    } else {
        # Sets the display name variable to last name + , + first name + initial.
        $tb_displayname.Text = $tb_lastname.Text + ", " + $tb_firstname.Text + " " + $tb_initial.Text + "."
    }
    
	#Update Account Name
    #If surname is not input yet the try sequence will fail and run the catch, which will continue with the first name only.
    try {
        #Sets the username variable to the first name + first letter of the last name, then enforced to be lowercase.
        #This could be improved by automatically adding more letters from the surname if the original username is taken in AD.
        $tb_accountname.Text = ($tb_firstname.Text -replace "'", '') + ($tb_lastname.Text.Substring(0,1) -replace "'", '')
        $tb_accountname.Text = $tb_accountname.Text.toLower()
    } catch {$tb_accountname.Text = $tb_firstname.Text.toLower() -replace "'", ''}

    #After the display name and username is updated, run functions to check the username validity and update the email address.
    AccountnameChanged
    EmailChanged
}

#This function will execute if the username is changed.
function AccountnameChanged {
    #If username does not match an existing username, display a green :) next to username field. Otherwise, if matching an existing username, display a red :(.
    if (!(Get-ADUser -Filter "sAMAccountName -eq '$($tb_accountname.Text)'")) {
        $label_validUsername.Text = ":)"
        $label_validUsername.ForeColor = 'green'
    } else {
        $label_validUsername.Text = ":("
        $label_validUsername.ForeColor = 'red'
    }
}

#This function will execute if the first name, last name, initial, domain name is changed, or if the account is set to be mail enabled.
function EmailChanged {
    # Update Email Address if email is enabled.
    # This function performs differently if initial is provided.
    if ($rb_emailaccYes.Checked -eq $true) {
        if ($tb_initial.Text -eq "") {
            # Sets the email address variable to first name + . + last name + domain. Removes any spaces and apostrophes (spaces cause email issues, apostrophes aren't compatible with eLearning's system for some reason)
            $tb_emailaddress.Text = $tb_firstname.text + "." + $tb_lastname.Text + $cb_emailtype.Text
            $tb_emailaddress.Text = $tb_emailaddress.Text -replace '\s', ''
            $tb_emailaddress.Text = $tb_emailaddress.Text -replace "'", ''
        } else {
            # Sets the email address variable to first name + . + initial + . + last name + domain. Removes any spaces and apostrophe (spaces cause email issues, apostrophes aren't compatible with eLearning's system for some reason).
            $tb_emailaddress.Text = $tb_firstname.text + "." + $tb_initial.Text + "." + $tb_lastname.Text + $cb_emailtype.Text
            $tb_emailaddress.Text = $tb_emailaddress.Text -replace '\s', ''
            $tb_emailaddress.Text = $tb_emailaddress.Text -replace "'", ''
        }
	}
}

#This function will execute if the OU is changed.
function OUChanged {
    # Changes the email address to the cell of the domain column in the same row of the OU in the CSV.
    # For example, if the OU is index/row 2 of the CSV, the domain will be set to index/row 2 of the CSV. 
    $cb_emailtype.Text = $email_array[$cb_ou.SelectedIndex]
}

#These variables execute the respective functions.
$TextChanged={
	TextChanged
}

$EmailChanged={
	EmailChanged
}

$AccountnameChanged={
    AccountnameChanged
}

$OUChanged={
    OUChanged
}

#This function will execute if the Mail Enabled radio buttons are changed.
$EmailAccChanged={
    # If mail enabled, enable the email field and set the domain according to the OU. (also see function OUChanged)
	if ($rb_emailaccYes.Checked -eq $true) {
		$cb_emailtype.Enabled = $true
        # If no OU selected, don't do anything.
        if ($cb_ou.SelectedIndex -ne 0) {
            $cb_emailtype.Text = $email_array[$cb_ou.SelectedIndex]
        }
        #After enabling the field and adjusting the domain, run a function to update the email address. 
		EmailChanged
	}
    #If mail is not enabled, disable the email field and blank the email address field.
	elseif ($rb_emailaccNo.Checked -eq $true) {
		$cb_emailtype.Enabled = $false
		$tb_emailaddress.Text = ""
	}
}

#=================================================================================================#
# MAIN FORM
[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(600,425)
$form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = "CenterScreen"
$form.Text = "$scriptVersion"
$form.Name = "$scriptVersion"
$form.Icon = $icon
#=================================================================================================#
# COMBO BOXES

# Email Type
$cb_emailtype = New-Object System.Windows.Forms.ComboBox
$cb_emailtype.Location = New-Object System.Drawing.Size(119,204)
$cb_emailtype.Size = New-Object System.Drawing.Size(200,21)
#When this field is changed, request the EmailChanged function.
$cb_emailtype.add_SelectedIndexChanged($EmailChanged)
#Import the CSV into an array to list.
$email_array = @()
Import-Csv $dbPath | ForEach-Object {$email_array += $_.Domain}
ForEach ($Item in $email_array | select -Unique) {[void] $cb_emailtype.Items.Add($Item)}
#Enabled by default.
$cb_emailtype.Enabled = $true

# OU Location
$cb_ou = New-Object System.Windows.Forms.ComboBox
$cb_ou.Location = New-Object System.Drawing.Size(119,181)
$cb_ou.Size = New-Object System.Drawing.Size(200,21)
#When this field is changed, request the OUChanged function.
$cb_ou.add_SelectedIndexChanged($OUChanged)
#Import the CSV into an array to list.
$ou_array = @()
Import-Csv $dBpath | ForEach-Object {$ou_array += $_.Unit}
ForEach ($Item in $ou_array) {[void] $cb_ou.Items.Add($Item)}
$cb_ou.TabIndex = 8;

#=================================================================================================#
# BUTTONS
$btn_ok = New-Object System.Windows.Forms.Button
$btn_ok.Text = "Create User"
$btn_ok.Location = New-Object System.Drawing.Size(475,330) 
$btn_ok.Size = New-Object System.Drawing.Size(80,40) 
$btn_ok.TabIndex = 12;

$btn_exit = New-Object System.Windows.Forms.Button
$btn_exit.Text = "EXIT"
$btn_exit.Location = New-Object System.Drawing.Size(25,330) 
$btn_exit.Size = New-Object System.Drawing.Size(80,40) 
$btn_exit.TabIndex = 14;

$btn_refreshpass = New-Object System.Windows.Forms.Button
$btn_refreshpass.Text = "R"
$btn_refreshpass.Location = New-Object System.Drawing.Size(272, 104) 
$btn_refreshpass.Size = New-Object System.Drawing.Size(20, 20)

#=================================================================================================#
# RADIO BUTTONS
$rb_emailaccNo = New-Object System.Windows.Forms.RadioButton
$rb_emailaccNo.Location = New-Object System.Drawing.Point(140,129)
$rb_emailaccNo.Size = New-Object System.Drawing.Size(50,22)
$rb_emailaccNo.Text = "No"
$rb_emailaccNo.add_CheckedChanged($EmailAccChanged)
$rb_emailaccNo.TabIndex = 5;

$rb_emailaccYes = New-Object System.Windows.Forms.RadioButton
$rb_emailaccYes.Location = New-Object System.Drawing.Point(200,129)
$rb_emailaccYes.Size = New-Object System.Drawing.Size(50,22)
$rb_emailaccYes.Text = "Yes"
$rb_emailaccYes.Checked = $true
$rb_emailaccYes.add_CheckedChanged($EmailAccChanged)
$rb_emailaccYes.TabIndex = 6;

#=================================================================================================#
# BUTTON ACTIONS

#This function will regenerate the password.
$btn_refreshpass.add_click({
    GeneratePassword
})

#This function will start when the OK button is pressed.
$btn_ok.add_click({
	
	# Checks if the first name, last name, username, or OU fields are not blank, otherwise error.
	if (($tb_firstname.Text -eq "") -or ($tb_lastname.Text -eq "") -or ($tb_accountname.Text -eq "") -or ($cb_ou.Text -eq "") -or ($tb_manager.Text -eq "")){
		Write-Host "Firstname, Lastname, Username, and OU field cannot be left blank." -ForegroundColor Red
		return
	}

    # To ensure mailbox licenses are costed accordingly, if costcode, request, or manager is not provided, error.
    if (($rb_emailaccYes.Checked -eq $true) -and (($tb_costcode.Text -eq "") -or ($tb_request.Text -eq "") -or ($tb_manager.Text -eq ""))) {
        Write-Host "A costcode, request number, and manager username must be provided for a mail enabled account." -ForegroundColor Red
        return
    }
	
	# Make sure the password is at least 14 characters long. This shouldn't generally happen since the password generator will always create an apt length password, but this is here for any manually input passwords.
	if ($tb_password.TextLength -lt 14) {
		Write-Warning "The entered password is less than 14 characters." -ForegroundColor Red
		return
	}
	
	# Make sure the account does not already exist. This shouldn't happen as the form indicator will let the analyst know the username is taken, but this code is here in case they try to create it anyway.
    #Create variables for the fields.
	$firstname = $tb_firstname.Text
	$accountname = $tb_accountname.Text
    #Create a variable and gets the user account in AD by their username. If no match, log result.
    try {$ADSearch = Get-ADUser $accountname} catch {Write-Host "No username conflicts found." -ForegroundColor DarkGray}
    #If the variable has a value from Get-ADUser, a conflict exists. 
	if ($ADSearch) {
		Write-Host
		Write-Warning "The account name $accountname has already been used.`nSee below for all users with a similar account" -ForegroundColor Red
        #Displays any alike usernames.
		$similarUsers = Get-ADUser -Filter "SamAccountName -like '$accountname*'" -Properties *| Select SamAccountName,DisplayName | FT
		$similarUsers | Out-Host
		Write-Host
		return
	}

	#================================================================================#
	# Determine combo box selections
	#================================================================================#
	# Email Domain

	if ($rb_emailaccYes.Checked -eq $true) {
	    $Script:emailDomain = $email_array[$cb_emailtype.SelectedIndex]
	}
	
	#================================================================================#	
	# OK to proceed. Return needed values from form to use in user creation.
	$Script:firstname = $tb_firstname.Text
    $Script:initial = $tb_initial.Text
	$Script:lastname = $tb_lastname.Text
	$Script:accountname = $tb_accountname.Text
	$Script:password = $tb_password.Text
    $Script:emailDomain = $cb_emailtype.Text
    $Script:costcode = $tb_costcode.Text
    $Script:email = $tb_emailaddress.Text
    $Script:ou = $cb_ou.Text
    $Script:display = $tb_displayname.Text
    $Script:request = $tb_request.Text
    $Script:manager = $tb_manager.Text

    # Close the form and continue running the script.
    $form.close()
})

#Exit button will escape out of the powershell script entirely.
$btn_exit.add_click({
	[environment]::exit(0)
})
#=================================================================================================#
# TEXTBOXES
$tb_firstname = New-Object System.Windows.Forms.TextBox 
$tb_firstname.Location = New-Object System.Drawing.Size(119,35) 
$tb_firstname.Size = New-Object System.Drawing.Size(150,10) 
$tb_firstname.add_TextChanged($TextChanged)
$tb_firstname.TabIndex = 1;
$tb_firstname.MaxLength = 28;

$tb_initial = New-Object System.Windows.Forms.TextBox
$tb_initial.Location = New-Object System.Drawing.Size(271,35)
$tb_initial.Size = New-Object System.Drawing.Size(30,10)
$tb_initial.add_TextChanged($TextChanged)

$tb_lastname = New-Object System.Windows.Forms.TextBox 
$tb_lastname.Location = New-Object System.Drawing.Size(119,58) 
$tb_lastname.Size = New-Object System.Drawing.Size(150,10)
$tb_lastname.add_TextChanged($TextChanged)
$tb_lastname.TabIndex = 2;
$tb_lastname.MaxLength = 28;

$tb_accountname = New-Object System.Windows.Forms.TextBox 
$tb_accountname.Location = New-Object System.Drawing.Size(119,81) 
$tb_accountname.Size = New-Object System.Drawing.Size(150,10) 
$tb_accountname.add_TextChanged($AccountnameChanged)
$tb_accountname.TabIndex = 3;
$tb_accountname.MaxLength = 20;

#This function will generate a new password
function GeneratePassword {
    #Blank slate to work on
    $ranPassword = ""
    #If the password is less than 15 characters, add another word.
    while ($ranPassword.length -le 15) {
        #A number is randomly generated between 1 and and the maximum number of lines in the Approved Words text file.
        $random = Get-Random -Maximum (Get-Content $wordsPath).length -Minimum 1
        #A row/line is then retrieved based on the given random number.
        $ranWord = Get-Content $wordsPath | select -Index $random
        #Random word is added to the Password variable.
        $ranPassword += $ranWord + " "
    }
    #If any trailing spaces are left, remove.
    $tb_password.Text = $ranPassword.Trim();
}

$tb_password = New-Object System.Windows.Forms.TextBox 
$tb_password.Location = New-Object System.Drawing.Size(119,104) 
$tb_password.Size = New-Object System.Drawing.Size(150,10) 
$tb_password.TabIndex = 4;
GeneratePassword

$tb_costcode = New-Object System.Windows.Forms.TextBox
$tb_costcode.Location = New-Object System.Drawing.Size(119,158)
$tb_costcode.Size = New-Object System.Drawing.Size(50,10)
$tb_costcode.TabIndex = 7;

$tb_displayname = New-Object System.Windows.Forms.TextBox 
$tb_displayname.Location = New-Object System.Drawing.Size(119,250) 
$tb_displayname.Size = New-Object System.Drawing.Size(150,10) 

$tb_emailaddress = New-Object System.Windows.Forms.TextBox 
$tb_emailaddress.Location = New-Object System.Drawing.Size(119,273)
$tb_emailaddress.Size = New-Object System.Drawing.Size(200,10)

$tb_request = New-Object System.Windows.Forms.TextBox 
$tb_request.Location = New-Object System.Drawing.Size(455,35) 
$tb_request.Size = New-Object System.Drawing.Size(100,10) 
$tb_request.TabIndex = 9;

$tb_manager = New-Object System.Windows.Forms.TextBox 
$tb_manager.Location = New-Object System.Drawing.Size(455,58) 
$tb_manager.Size = New-Object System.Drawing.Size(100,10) 
$tb_manager.TabIndex = 10;

$tb_displayname.ReadOnly = $true
$tb_emailaddress.ReadOnly = $true

#=================================================================================================#
# LABELS
$label_header = New-Object Windows.Forms.label
$label_header.Text = "New User Creation"

$label_header.AutoSize = $false
$label_header.TextAlign = "MiddleCenter"
$label_header.Dock = "Top"
$label_header.Font = $font_header

$label_firstname = New-Object Windows.Forms.label
$label_firstname.Text = "First Name"
$label_firstname.AutoSize = $false
$label_firstname.TextAlign = "MiddleLeft"
$label_firstname.Location = New-Object System.Drawing.Size(12,33) 

$label_lastname = New-Object Windows.Forms.label
$label_lastname.Text = "Last Name"
$label_lastname.AutoSize = $false
$label_lastname.TextAlign = "MiddleLeft"
$label_lastname.Location = New-Object System.Drawing.Size(12,56) 

$label_accountname = New-Object Windows.Forms.label
$label_accountname.Text = "Username"
$label_accountname.AutoSize = $false
$label_accountname.TextAlign = "MiddleLeft"
$label_accountname.Location = New-Object System.Drawing.Size(12,79) 

$label_validUsername = New-Object Windows.Forms.label
$label_validUsername.Text = " :("
$label_validUsername.AutoSize = $false
$label_validUsername.ForeColor = 'red'
$label_validUsername.width = 40
$label_validUsername.TextAlign = "MiddleLeft"
$label_validUsername.Location = New-Object System.Drawing.Size(270,79) 

$label_password = New-Object Windows.Forms.label
$label_password.Text = "Password"
$label_password.AutoSize = $false
$label_password.TextAlign = "MiddleLeft"
$label_password.Location = New-Object System.Drawing.Size(12,102) 

$label_emailacc = New-Object Windows.Forms.label
$label_emailacc.Text = "Mail Enabled User?"
$label_emailacc.AutoSize = $false
$label_emailacc.TextAlign = "MiddleLeft"
$label_emailacc.Location = New-Object System.Drawing.Size(12,132)
$label_emailacc.Size = New-Object System.Drawing.Size(120,15)

$label_costcode = New-Object Windows.Forms.label
$label_costcode.Text = "Costcode"
$label_costcode.AutoSize = $false
$label_costcode.TextAlign = "MiddleLeft"
$label_costcode.Location = New-Object System.Drawing.Size(12,158)

$label_ou = New-Object Windows.Forms.label
$label_ou.Text = "OU"
$label_ou.AutoSize = $false
$label_ou.TextAlign = "MiddleLeft"
$label_ou.Location = New-Object System.Drawing.Size(12,181) 

$label_emailtype = New-Object Windows.Forms.label
$label_emailtype.Text = "Email Domain"
$label_emailtype.AutoSize = $false
$label_emailtype.TextAlign = "MiddleLeft"
$label_emailtype.Location = New-Object System.Drawing.Size(12,204) 

$label_displayname = New-Object Windows.Forms.label
$label_displayname.Text = "Display Name"
$label_displayname.AutoSize = $false
$label_displayname.TextAlign = "MiddleLeft"
$label_displayname.Location = New-Object System.Drawing.Size(12,248) 

$label_emailaddress = New-Object Windows.Forms.label
$label_emailaddress.Text = "Mail/SIP Address"
$label_emailaddress.AutoSize = $false
$label_emailaddress.TextAlign = "MiddleLeft"
$label_emailaddress.Location = New-Object System.Drawing.Size(12,271) 

$label_request = New-Object Windows.Forms.label
$label_request.Text = "Ticket Request #"
$label_request.AutoSize = $false
$label_request.TextAlign = "MiddleLeft"
$label_request.Location = New-Object System.Drawing.Size(330,34) 

$label_manager = New-Object Windows.Forms.label
$label_manager.Text = "Manager's Username"
$label_manager.AutoSize = $false
$label_manager.Location = New-Object System.Drawing.Size(330,60) 
$label_manager.TextAlign = "MiddleLeft"
$label_manager.Size = new-object System.Drawing.Size(120,15)

$image_path = $logoPath
$image = [System.Drawing.Image]::FromFile("$image_path")
$image_logo = New-Object Windows.Forms.PictureBox
$image_logo.Location = New-Object System.Drawing.Size(340,120)
$image_logo.Width = $image.Size.Width
$image_logo.Height = $image.Size.Height
$image_logo.Image = $image

#I used to have a section here to show available E3 licenses, but it was removed as it was painful having to input an AAD password every time I wanted to create a new user.
#If ever needing the functionality, code to retrieve the license numbers are below.

<#
$availDefenderPlans = Get-MsolAccountSku | where {$_.AccountSkuId -eq "INPUTIDHERE"} | Select -ExpandProperty "ActiveUnits"
$availExchangePlans = Get-MsolAccountSku | where {$_.AccountSkuId -eq "INPUTIDHERE"} | Select -ExpandProperty "ActiveUnits"
#>

#=================================================================================================#
# ADD ELEMENTS TO FORM
$form.Controls.Add($label_header)
$form.Controls.Add($label_firstname)
$form.Controls.Add($label_lastname)
$form.Controls.Add($label_accountname)
$form.Controls.Add($label_password)
$form.Controls.Add($label_emailacc)
$form.Controls.Add($label_emailtype)
$form.Controls.Add($label_displayname)
$form.Controls.Add($label_emailaddress)
$form.Controls.Add($label_ou)
$form.Controls.Add($label_costcode)
$form.Controls.Add($label_request)
$form.Controls.Add($label_manager)
$form.Controls.Add($label_validUsername)

$form.Controls.Add($tb_firstname) 
$form.Controls.Add($tb_initial)
$form.Controls.Add($tb_request)

$form.Controls.Add($tb_lastname) 
$form.Controls.Add($tb_accountname) 
$form.Controls.Add($tb_password) 
$form.Controls.Add($tb_displayname) 
$form.Controls.Add($tb_emailaddress)
$form.Controls.Add($tb_costcode)
$form.Controls.Add($tb_manager)
$form.Controls.Add($image_logo)

$form.Controls.Add($btn_ok)
$form.Controls.Add($btn_exit)
$form.Controls.Add($btn_refreshpass)

$form.Controls.Add($rb_emailaccYes)
$form.Controls.Add($rb_emailaccNo)

$form.Controls.Add($cb_emailtype)
$form.Controls.Add($cb_ou)
#=================================================================================================#


#==================#
#===  END FORM  ===#
#==================#

$form.Add_Shown({$form.Activate()})
[void] $form.ShowDialog()


#=================================================================================================#
# Script Begins
#=================================================================================================#
CLEAR

Write-Host "`nAttempting to create user: $accountname. Please wait..." -Fore White
SLEEP 2

#=================================================================================================#
# User Creation
#=================================================================================================#

#Imports data from the CSV row depending on the OU that is selected.
$importData = import-csv $dbPath | Where-Object {$_."Unit" -eq "$ou"}

#Creates variables that imports the relevant fields from the CSV, based on the column headers.
$description = $importData."Description"
$office = $importData."Office"
$loginscript = $importData."Loginscript"
$servername = $importData."HomeServer"
$department = $importData."Department"
$oupath = $importData."OU"
$homedir = "\\$($servername)\$($accountname)" + "$"
$homedirpath = "\\$($servername)\users$\$($accountname)"
$sharename = "$($accountname)" + "$"
$sharedirpath = "e:\data\users\$($accountname)"
$departmentnumber = $department + "-" + "$costcode"

#This script stops itself on purpose to not accidentally create anything right now
[environment]::exit(0)

#Creates the user.
try {
New-ADUser -Name $display -SamAccountName $accountname -DisplayName $display `
-givenname $firstname -surname $lastname -initials $initial -userprincipalname $email `
-Path $oupath -Enabled $true -ChangePasswordAtLogon $true -Department $department `
-OtherAttributes @{'departmentNumber'="$departmentnumber"} -HomeDrive "M" -HomeDirectory $homedir `
-Description $description -Office $office -ScriptPath $loginscript -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force)

Write-Host "Created user." -ForegroundColor DarkGray
} catch {
    #If any failures with account creation, exit the powershell script.
    [environment]::exit(0)
}

#Creates variable to start at column 1 "Group1" of the CSV.
$a = 1
$GroupID = "Groups"+$a
#If that column cell of the relevant row/OU is not empty, continue.
while ($importData.$GroupID -ne $null) {
    Add-ADGroupMember -Identity $importData.$GroupID -Members $accountname
    #Go to the next column, and according to the while loop, ensure it's not empty.
    $a++
    $GroupID = "Groups"+$a
}
Write-Host "Added AD groups." -ForegroundColor DarkGray

#Creates the M drive and sets permissions.
New-Item -path $homedirpath -type directory
$acl = Get-ACL -path $homedirpath
$permission = "DOMAIN\$($accountname)","Modify","ContainerInherit,ObjectInherit","None","Allow"
$accessrule = new-object System.Security.AccessControl.FileSystemAccessRule $permission
$acl.SetAccessRule($accessrule)
$acl | Set-ACL -path $homedirpath

#Sets up more M drive stuff. I think Damon created this section so I have no idea what it does step by step.
$Computer = $servername
$Class = "Win32_Share"
$Method = "Create"
$name = $sharename
$path = $sharedirpath
$description = ""
$sd = ([WMIClass] "\\$Computer\root\cimv2:Win32_SecurityDescriptor").CreateInstance()
$ACE = ([WMIClass] "\\$Computer\root\cimv2:Win32_ACE").CreateInstance()
$Trustee = ([WMIClass] "\\$Computer\root\cimv2:Win32_Trustee").CreateInstance()
$Trustee.Name = "EVERYONE"
$Trustee.Domain = $Null
$Trustee.SID = @(1, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0)
$ace.AccessMask = 2032127
$ace.AceFlags = 3
$ace.AceType = 0
$ACE.Trustee = $Trustee
$sd.DACL += $ACE.psObject.baseobject 
$mc = [WmiClass]"\\$Computer\ROOT\CIMV2:$Class"
$InParams = $mc.psbase.GetMethodParameters($Method)
$InParams.Access = $sd
$InParams.Description = $description
$InParams.MaximumAllowed = $Null
$InParams.Name = $name
$InParams.Password = $Null
$InParams.Path = $path
$InParams.Type = [uint32]0
$R = $mc.PSBase.InvokeMethod($Method, $InParams, $Null)
switch ($($R.ReturnValue))
 {
  0 {Write-Host "Share:$name Path:$path Result:Success"; break}
  2 {Write-Host "Share:$name Path:$path Result:Access Denied" -foregroundcolor red -backgroundcolor yellow;break}
  8 {Write-Host "Share:$name Path:$path Result:Unknown Failure" -foregroundcolor red -backgroundcolor yellow;break}
  9 {Write-Host "Share:$name Path:$path Result:Invalid Name" -foregroundcolor red -backgroundcolor yellow;break}
  10 {Write-Host "Share:$name Path:$path Result:Invalid Level" -foregroundcolor red -backgroundcolor yellow;break}
  21 {Write-Host "Share:$name Path:$path Result:Invalid Parameter" -foregroundcolor red -backgroundcolor yellow;break}
  22 {Write-Host "Share:$name Path:$path Result:Duplicate Share" -foregroundcolor red -backgroundcolor yellow;break}
  23 {Write-Host "Share:$name Path:$path Result:Reedirected Path" -foregroundcolor red -backgroundcolor yellow;break}
  24 {Write-Host "Share:$name Path:$path Result:Unknown Device or Directory" -foregroundcolor red -backgroundcolor yellow;break}
  25 {Write-Host "Share:$name Path:$path Result:Network Name Not Found" -foregroundcolor red -backgroundcolor yellow;break}
  default {Write-Host " " -ForegroundColor DarkGray}
 }
Write-Host "Created and shared M drive." -ForegroundColor DarkGray

#Start email process if the email address variable exists.
if ($email) {
    Write-Host "User account has been created. In 15 minutes you'll be reminded to create the Exchange mailbox." -ForegroundColor White
    write-host ("15 minutes remaining.") #0
    start-sleep (60*1); write-host ("14 minutes remaining.") #1
    start-sleep (60*1); write-host ("13 minutes remaining.") #2
    start-sleep (60*1); write-host ("12 minutes remaining.") #3
    start-sleep (60*1); write-host ("11 minutes remaining.") #4
    start-sleep (60*1); write-host ("10 minutes remaining.") #5
    start-sleep (60*1); write-host ("9 minutes remaining.") #6
    start-sleep (60*1); write-host ("8 minutes remaining.") #7
    start-sleep (60*1); write-host ("7 minutes remaining.") #8
    start-sleep (60*1); write-host ("6 minutes remaining.") #9
    start-sleep (60*1); write-host ("5 minutes remaining.") #10
    start-sleep (60*1); write-host ("4 minutes remaining.") #11
    start-sleep (60*1); write-host ("3 minutes remaining.") #12
    start-sleep (60*1); write-host ("2 minutes remaining.") #13
    start-sleep (60*1); write-host ("1 minutes remaining.") #14
    start-sleep (60*1); write-host ("`a"*4)

    [System.Windows.Forms.MessageBox]::Show("Please create the Exchange mailbox now.
    Click OK to receive the command. Then close the window.
    You'll be notified when the mailbox linking has finalised.","$scriptVersion")

    #Generates a form to provide the command for Exchange server.
    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object Drawing.Size @(400,100)
    $form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.StartPosition = "CenterScreen"
    $form.Text = "$scriptVersion"
    $form.Name = "$scriptVersion"
    $form.Icon = $icon

    $label_exchCode = New-Object Windows.Forms.label
    $label_exchCode.Text = "Input Into Exchange Management Shell"
    $label_exchCode.Location = New-Object System.Drawing.Size(25,15)
    $label_exchCode.Size = New-Object System.Drawing.Size(350,15) 

    $tb_code = New-Object System.Windows.Forms.TextBox 
    $tb_code.Location = New-Object System.Drawing.Size(20,35) 
    $tb_code.Size = New-Object System.Drawing.Size(300,10) 

    #This copy button doesn't work sometimes and it's not my fault. Windows form copies are just really dicky over remote connection/server.
    $btn_copy = New-Object System.Windows.Forms.Button
    $btn_copy.Text = "Copy"
    $btn_copy.Location = New-Object System.Drawing.Size(325,35) 
    $btn_copy.Size = New-Object System.Drawing.Size(40,20) 
    
    #When copy button is pressed, add command to the clipboard.
    $btn_copy.add_click({
        Set-Clipboard -Value $tb_code.Text
    })

    $tb_code.ReadOnly = $true

    $form.Controls.Add($label_exchCode)
    $form.Controls.Add($btn_copy)
    $form.Controls.Add($tb_code)

    #Generate a command to be run on the Exchange server for mailbox creation.
    $tb_code.Text = 'Enable-RemoteMailbox -Identity ' + '"' + $display + '"' + ' -RemoteRoutingAddress ' + '"' + $accountname + $onmicrosoftDomain + '"'

    $form.Add_Shown({$form.Activate()})
    [void] $form.ShowDialog()


    #Once Exchange has synced, it populates the 'mail' attribute.
    #Once this change is detected, it will continue the script.
    #Attributes can only be changed after the sync otherwise Exchange will overwrite the changes made.
    while($(Get-ADUser $accountname -Properties mail).mail -eq $null){
        start-sleep (30)
        Write-host "Sync incomplete."
    }
    Write-Host "Sync Complete.
    Found Entries and Replaced."
    
    #Convert from the default email domain to the custom domain, as set by the CSV entry.
    Set-ADUser $accountname -replace @{mail=$email}
    Set-AdUser $accountname -remove @{ProxyAddresses="SMTP:"+$firstname+"."+$lastname+$defaultDomain}
    Set-AdUser $accountname -remove @{ProxyAddresses="SMTP:"+$firstname+"."+$initial+"."+$lastname+$defaultDomain}
    Set-AdUser $accountname -add @{ProxyAddresses="SMTP:"+$email}

    #Below prevents Exchange settings syncing and overwriting the AD account's setting. 
    #Which allows the change of the email domain and also any name changes for future cases.
    Set-ADUser $accountname -remove @{msExchPoliciesIncluded="{26491cfc-9e50-4857-861b-0cb8df22b5d7}"}
    Set-ADUser $accountname -add @{msExchPoliciesExcluded="{26491cfc-9e50-4857-861b-0cb8df22b5d7}"}
    
    Write-Host "Set proxy addresses." -ForegroundColor DarkGray

    #Add groups that apply to all users
    Add-ADGroupMember -Identity "All Users" -Members $accountname

    Write-Host "Assigned default groups." -ForegroundColor DarkGray
} else {
    #If account is not mail enabled, display success notice.
    [System.Windows.Forms.MessageBox]::Show("User account created!","$scriptVersion")
}

#=================================================================================================#
# Form Detailing User Info for LANDesk
#=================================================================================================#

#End form to display user information.
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(300,200)
$form.FormBorderStyle=[System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = "CenterScreen"
$form.Text = "$scriptVersion"
$form.Name = "$scriptVersion"
$form.Icon = $icon

$label_header = New-Object Windows.Forms.label
$label_header.Text = "New User Creation"
$label_header.AutoSize = $false
$label_header.TextAlign = "MiddleCenter"
$label_header.Dock = "Top"
$label_header.Font = $font_header

$label_username = New-Object Windows.Forms.label
$label_username.Text = "Username:"
$label_username.AutoSize = $false
$label_username.TextAlign = "MiddleLeft"
$label_username.Location = New-Object System.Drawing.Size(12,33) 

$label_email = New-Object Windows.Forms.label
$label_email.Text = "Email Address:"
$label_email.AutoSize = $false
$label_email.TextAlign = "MiddleLeft"
$label_email.Location = New-Object System.Drawing.Size(12,56) 

$label_pass = New-Object Windows.Forms.label
$label_pass.Text = "Password:"
$label_pass.AutoSize = $false
$label_pass.TextAlign = "MiddleLeft"
$label_pass.Location = New-Object System.Drawing.Size(12,79) 

#Remaining things for the analyst to do that the script cannot do.
$label_other = New-Object Windows.Forms.label
$label_other.Text = "- Send Notification to Manager
- Resolve Ticket"
$label_other.Location = New-Object System.Drawing.Size(12,115) 
$label_other.Size = New-Object System.Drawing.Size(200,50) 

$tb_username = New-Object System.Windows.Forms.TextBox 
$tb_username.Location = New-Object System.Drawing.Size(115,35) 
$tb_username.Size = New-Object System.Drawing.Size(150,10) 

$tb_email = New-Object System.Windows.Forms.TextBox 
$tb_email.Location = New-Object System.Drawing.Size(115,58)
$tb_email.Size = New-Object System.Drawing.Size(150,10)

$tb_pass = New-Object System.Windows.Forms.TextBox 
$tb_pass.Location = New-Object System.Drawing.Size(115,81)
$tb_pass.Size = New-Object System.Drawing.Size(150,10)

$tb_username.ReadOnly = $true
$tb_email.ReadOnly = $true
$tb_pass.ReadOnly = $true

$form.Controls.Add($label_header)
$form.Controls.Add($label_username)
$form.Controls.Add($label_email)
$form.Controls.Add($label_pass)
$form.Controls.Add($label_other)
$form.Controls.Add($tb_username)
$form.Controls.Add($tb_email)
$form.Controls.Add($tb_pass)

$tb_username.Text = $accountname
$tb_email.Text = $email
$tb_pass.Text = $password

$form.Add_Shown({$form.Activate()})
[void] $form.ShowDialog()
