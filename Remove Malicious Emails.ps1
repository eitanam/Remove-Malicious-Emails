<#

***************   This script is provided AS-IS without any warranty to any damage that may occured. It will delete emails, if you are using it it's AT YOUR OWN RISK!  ***********
Version 3.1
Removed the option to connect without MFA
Updated the connection to V3

Version 3.0

Improved GUI
Added an option to search by sender adreess and date range
Fixed some minor bugs

Version 2.0

As Microsoft changed the search way I re-write the code. The code is now using the Office 365 Security & Compliance

Version 1.1
Check if the Exchange Online PowerShell using multi-factor authentication module is installed

Version 1.0
Inital release


#>

$path = New-Item -ItemType Directory -Force -Path C:\temp1
$LogFile = '\AffectedMailboxes.csv'
$PurgeFile = '\PurgeResults.csv'
$LogFullPath = Join-Path -path $path -ChildPath $LogFile
$PurgeFullPath = Join-Path -path $path -ChildPath $PurgeFile

$str001 = "Remove Malicious Emails Ver 3.1"
$str002 = "***********    Please make sure you are a member of the eDiscovery Manager role and the Organization Management groups    **********"
$str003 = "To check if you are a member of one of those groups go to https://protection.office.com and on the left pane click  Permissions"
$str004 = "I will search in your entire environment, but I will return only a maximum of 500 Mailboxes in the results."
$str005 = "The Purge Action will purge a maximum of 10 Items per Mailbox."
$str010 = "Connection Status:"
$str012 = "Type a Name for this search (less than 198 Characters)"
$str013 = "Sender Email Address"
$str014 = "Email Subject"
$str015 = "Days To Search"
$str016 = "Search name"
$str017 = "Recipient email address? (to search in all MB's type all)"
$str018 = "How to search?"
$str019 = "By Subject, Sender Address and Date Range"
$str020 = "By Subject and Date Range"
$str021 = "By Sender Address and Date Range"
$str022 = "By Subject and Sender Address"
$str023 = "Where and what to search?"
$str024 = "Actions"
$str025 = "Search"
$str026 = "Get a list of the affected mailboxes"
$str027 = "Delete the emails"
$str028 = "Info"
$str029 = "Status"
$str030 = "I Created the search, Now I am starting it"
$str031 = "I am generating the report, please wait"
$str035 = "Connected to Security and Compliance Center"
$str036 = "Some fileds are missing data"
$str037 = "Red"
$str038 = "Green"
$str039 = "IndianRed"
$str040 = "Search has been completed"
$str041 = "Close"



Function Connect()
{
	$exoerr = "clean"
	$EXOcheck = Get-InstalledModule ExchangeOnlineManagement
	if ($EXOcheck.Version.ToString() -match "^[0-2](\.[0-9]+){0,2}$")
	{
		Update-Module -Name ExchangeOnlineManagement
	}
	try
	{
		Import-Module ExchangeOnlineManagement -EA Stop
	}
	catch
	{
	    Install-Module ExchangeOnlineManagement
	}
	try
	{
    	Connect-IPPSSession #-ErrorAction Stop -ErrorVariable exoerr
	}
	catch
	{
       	if ($exoerr -match "User canceled authentication.")
		{
           	[System.Windows.Forms.MessageBox]::Show("User Canceled Authentication, Please Try again", "Error", 0, 
           	[System.Windows.Forms.MessageBoxIcon]::Exclamation)
			Connect
		}
        else
		{
            [System.Windows.Forms.MessageBox]::Show("Could not connect to Exchange Online Powershell, Error: $exoerr", "Error", 0, 
            [System.Windows.Forms.MessageBoxIcon]::Exclamation)
            Exit 1
        }
    }
}

Connect

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 
$Form                                      = New-Object system.Windows.Forms.Form
$Form.ClientSize                           = '900,700'
$Form.text                                 = $str001
$Form.TopMost                              = $false
#
$StatusLabel                               = New-Object system.Windows.Forms.Label
$StatusLabel.text                          = $str010
$StatusLabel.AutoSize                      = $true
$StatusLabel.location                      = New-Object System.Drawing.Point(280,110)
$StatusLabel.Font                          = 'Microsoft Sans Serif,10'
#
$StatusUpdate                              = New-Object system.Windows.Forms.Label
$StatusUpdate.ForeColor                    = $str038
$StatusUpdate.Text                         = $str035
$StatusUpdate.AutoSize                     = $true
$StatusUpdate.location                     = New-Object System.Drawing.Point(350,50)
$StatusUpdate.Font                         = 'Microsoft Sans Serif,10'
#
#
$SearchGroup                               = New-Object 'System.Windows.Forms.GroupBox'
$SearchName                                = New-Object 'system.Windows.Forms.Label'
$SearchNameTextBox                         = New-Object 'system.Windows.Forms.TextBox'
#
$SearchGroup.Location                      = New-Object System.Drawing.Point(30, 93)
$SearchGroup.Size                          = New-Object System.Drawing.Size(825, 50)
$SearchGroup.Text                          = $str016
$SearchGroup.TabIndex                      = 0
$SearchGroup.TabStop                       = $False
#
$SearchName.text                           = $str012
$SearchName.AutoSize                       = $true
$SearchName.location                       = New-Object System.Drawing.Point(40,110)
$SearchName.font                           = 'Microsoft Sans Serif,10'
#
$SearchNameTextBox.multiline               = $false
$SearchNameTextBox.Size                    = New-Object System.Drawing.Size(125,250)
$SearchNameTextBox.location                = New-Object System.Drawing.Point(390,110)
$SearchNameTextBox.TabIndex                = 0
#
#
$HowToSearchGroup                          = New-Object 'System.Windows.Forms.GroupBox'
$radiobutton1                              = New-Object 'System.Windows.Forms.RadioButton'
$radiobutton2                              = New-Object 'System.Windows.Forms.RadioButton'
$radiobutton3                              = New-Object 'System.Windows.Forms.RadioButton'
$radiobutton4                              = New-Object 'System.Windows.Forms.RadioButton'
#
$HowToSearchGroup.Location                 = New-Object System.Drawing.Point(30, 153)
$HowToSearchGroup.Size                     = New-Object System.Drawing.Size(320, 160)
$HowToSearchGroup.Text                     = $str018
$HowToSearchGroup.TabIndex                 = 1
$HowToSearchGroup.TabStop                  = $False
#
$radiobutton1.Location                     = New-Object System.Drawing.Point(40,180)
$radiobutton1.AutoSize                     = $true
$radiobutton1.Text                         = $str019 
$radiobutton1.UseCompatibleTextRendering   = $True
$radiobutton1.UseVisualStyleBackColor      = $True
$radiobutton1.Checked                      = $false
$radiobutton1.Font                         = 'Microsoft Sans Serif,10'
$radiobutton1.TabIndex                     = 0
$radiobutton1.ADD_CheckedChanged({radiobutton_CheckedChanged})
#
$radiobutton2.Location                     = New-Object System.Drawing.Point(40, 210)
$radiobutton2.AutoSize                     = $true
$radiobutton2.Text                         = $str020 
$radiobutton2.UseCompatibleTextRendering   = $True
$radiobutton2.UseVisualStyleBackColor      = $True
$radiobutton2.Checked                      = $false
$radiobutton2.Font                         = 'Microsoft Sans Serif,10'
$radiobutton2.TabIndex                     = 1
$radiobutton2.ADD_CheckedChanged({radiobutton_CheckedChanged})
#
$radiobutton3                              = New-Object 'System.Windows.Forms.RadioButton'
$radiobutton3.Location                     = New-Object System.Drawing.Point(40, 240)
$radiobutton3.AutoSize                     = $true
$radiobutton3.Text                         = $str021
$radiobutton3.UseCompatibleTextRendering   = $True
$radiobutton3.UseVisualStyleBackColor      = $True
$radiobutton3.Checked                      = $false
$radiobutton3.Font                         = 'Microsoft Sans Serif,10'
$radiobutton3.TabIndex                     = 2
$radiobutton3.ADD_CheckedChanged({radiobutton_CheckedChanged})
#
$radiobutton4.Location                     = New-Object System.Drawing.Point(40, 270)
$radiobutton4.AutoSize                     = $true
$radiobutton4.Text                         = $str022
$radiobutton4.UseCompatibleTextRendering   = $True
$radiobutton4.UseVisualStyleBackColor      = $True
$radiobutton4.Checked                      = $false
$radiobutton4.Font                         = 'Microsoft Sans Serif,10'
$radiobutton4.TabIndex                     = 3
$radiobutton4.ADD_CheckedChanged({radiobutton_CheckedChanged})
#
#
$QueryDetailsGroup                         = New-Object 'System.Windows.Forms.GroupBox'
$SearchLocationLabel                       = New-Object 'system.Windows.Forms.Label'
$SearchLocationTxtbox                      = New-Object 'system.Windows.Forms.TextBox'
$EmailAddress                              = New-Object 'system.Windows.Forms.Label'
$EmailAddressTextBox                       = New-Object 'system.Windows.Forms.TextBox'
$Subject                                   = New-Object 'system.Windows.Forms.Label'
$SubjectTxtbox                             = New-Object 'system.Windows.Forms.TextBox'
$Days                                      = New-Object 'system.Windows.Forms.Label'
$DaysTextBox                               = New-Object 'system.Windows.Forms.TextBox'
#
$QueryDetailsGroup.Location                = New-Object System.Drawing.Point(355, 153)
$QueryDetailsGroup.Size                    = New-Object System.Drawing.Size(500, 160)
$QueryDetailsGroup.Text                    = $str023
$QueryDetailsGroup.TabIndex                = 2
$QueryDetailsGroup.TabStop                 = $False
#
$SearchLocationLabel.text                  = $str017
$SearchLocationLabel.AutoSize              = $true
$SearchLocationLabel.location              = New-Object System.Drawing.Point(365,180)
$SearchLocationLabel.Font                  = 'Microsoft Sans Serif,10'
#
$SearchLocationTxtbox.Location             = New-Object System.Drawing.Point(710,180)
$SearchLocationTxtbox.Size                 = New-Object System.Drawing.Size(130,20)
$SearchLocationTxtbox.TabIndex             = 0
$SearchLocationTxtbox.TabStop              = $False
$SearchLocationTxtbox.Enabled              = $false
#
$EmailAddress.text                         = $str013
$EmailAddress.AutoSize                     = $true
$EmailAddress.location                     = New-Object System.Drawing.Point(365,210)
$EmailAddress.font                         = 'Microsoft Sans Serif,10'
$EmailAddress.Enabled                      = $false
#
$EmailAddressTextBox                       = New-Object system.Windows.Forms.TextBox
$EmailAddressTextBox.multiline             = $false
$EmailAddressTextBox.Size                  = New-Object System.Drawing.Point(330,50)
$EmailAddressTextBox.location              = New-Object System.Drawing.Point(510,210)
$EmailAddressTextBox.Font                  = 'Microsoft Sans Serif,10'
$EmailAddressTextBox.Enabled               = $false
$EmailAddressTextBox.TabIndex              = 1
#
$Subject.text                              = $str014
$Subject.AutoSize                          = $true
$Subject.location                          = New-Object System.Drawing.Point(365,240)
$Subject.Font                              = 'Microsoft Sans Serif,10'
#
$SubjectTxtbox.multiline                   = $false
$SubjectTxtbox.size                        = New-Object System.Drawing.Point(330,50)
$SubjectTxtbox.location                    = New-Object System.Drawing.Point(510,240)
$SubjectTxtbox.Font                        = 'Microsoft Sans Serif,10'
$SubjectTxtbox.Enabled                     = $false
$SubjectTxtbox.TabIndex                    = 2
#
$Days.text                                 = $str015
$Days.AutoSize                             = $true
$Days.location                             = New-Object System.Drawing.Point(365,270)
$Days.Font                                 = 'Microsoft Sans Serif,10'
#
$DaysTextBox.multiline                     = $false
$DaysTextBox.size                          = New-Object System.Drawing.Point(100,50)
$DaysTextBox.location                      = New-Object System.Drawing.Point(510,270)
$DaysTextBox.Font                          = 'Microsoft Sans Serif,10'
$DaysTextBox.Enabled                       = $false
$DaysTextBox.TabIndex                      = 3
#
#
$ActionsGroup                              = New-Object 'System.Windows.Forms.GroupBox'
$SearchButton                              = New-Object 'system.Windows.Forms.Button'
$LogsButton                                = New-Object 'system.Windows.Forms.Button'
$DeleteButton                              = New-Object 'system.Windows.Forms.Button'
#
$ActionsGroup.Location                     = New-Object System.Drawing.Point(30,320)
$ActionsGroup.Size                         = New-Object System.Drawing.Size(320,170)
$ActionsGroup.Text                         = $str024
$ActionsGroup.TabIndex                     = 3
$ActionsGroup.TabStop                      = $False
#
#
$SearchButton.text                         = $str025
$SearchButton.autosize                     = $true
$SearchButton.location                     = New-Object System.Drawing.Point(60,350)
$SearchButton.Font                         = 'Microsoft Sans Serif,10'
$SearchButton.visible                      = $True
#
$LogsButton.text                           = $str026
$LogsButton.autosize                       = $true
$LogsButton.location                       = New-Object System.Drawing.Point(60,390)
$LogsButton.Font                           = 'Microsoft Sans Serif,10'
$LogsButton.visible                        = $false
#
$DeleteButton.text                         = $str027
$DeleteButton.autosize                     = $True
$DeleteButton.location                     = New-Object System.Drawing.Point(60,430)
$DeleteButton.Font                         = 'Microsoft Sans Serif,10'
$DeleteButton.visible                      = $false
#
#
$ResultsGroup                              = New-Object 'System.Windows.Forms.GroupBox'
$ResultTxtbox                              = New-Object 'system.Windows.Forms.TextBox'
#
$ResultsGroup.Location                     = New-Object System.Drawing.Point(355,320)
$ResultsGroup.Size                         = New-Object System.Drawing.Size(500,280)
$ResultsGroup.Text                         = $str028
$ResultsGroup.TabIndex                     = 4
$ResultsGroup.TabStop                      = $False
#
$ResultTxtbox.multiline                    = $true
$ResultTxtbox.size                         = New-Object System.Drawing.Point(490,250)
$ResultTxtbox.location                     = New-Object System.Drawing.Point(360,340)
$ResultTxtbox.BackColor                    = [System.Drawing.Color]::FromArgb(245,245,220) #, 192)
$ResultTxtbox.text                         = "                                   Welcome to "+ $str001 + "`r`n `r`n" +$str002 + "`r`n" +$str003 + "`r`n`r`n" +$str004 + "." +$str005 + "`r`n`r`n" +$str006 + "`r`n"
$ResultTxtbox.Font                         = 'Microsoft Sans Serif,10'
#
$closeButton                               = New-Object system.Windows.Forms.Button
$closeButton.text                          = $str041
$closeButton.autosize                      = $True
$closeButton.location                      = New-Object System.Drawing.Point(750,620)
#
#
$SearchResultsGroup                        = New-Object 'System.Windows.Forms.GroupBox'
#
$SearchResultsGroup.Location               = New-Object System.Drawing.Point(30,490)
$SearchResultsGroup.Size                   = New-Object System.Drawing.Size(320,110)
$SearchResultsGroup.Text                   = $str029
$SearchResultsGroup.TabIndex               = 5
$SearchResultsGroup.TabStop                = $False
#
$SearchStatus                              = New-Object system.Windows.Forms.Label
$SearchStatus.AutoSize                     = $true
$SearchStatus.location                     = New-Object System.Drawing.Point(40,530)
$SearchStatus.Font                         = 'Microsoft Sans Serif,10'
$form.Controls.Add($SearchStatus)
$form.controls.add($StatusUpdate)
$form.Controls.Add($SearchName)
$form.Controls.Add($SearchNameTextBox)
$form.Controls.Add($radiobutton1)
$form.Controls.Add($radiobutton2)
$form.Controls.Add($radiobutton3)
$form.Controls.Add($radiobutton4)
$form.Controls.Add($SearchGroup)
$form.Controls.Add($HowToSearchGroup)
$form.Controls.Add($QueryDetailsGroup)
$form.Controls.Add($SearchLocationLabel)
$form.Controls.Add($SearchLocationTxtbox)
$form.Controls.Add($EmailAddress)
$form.Controls.Add($EmailAddressTextBox)
$form.Controls.Add($Subject)
$form.Controls.Add($SubjectTxtbox)
$form.Controls.Add($Days)
$form.Controls.Add($DaysTextBox)
$form.Controls.Add($SearchButton)
$form.Controls.Add($LogsButton)
$form.Controls.Add($DeleteButton)
$form.Controls.Add($ResultTxtbox)
$form.Controls.Add($ActionsGroup)   
$form.Controls.Add($ResultsGroup)
$form.Controls.Add($QueryDetailsGroup)
$form.Controls.Add($SearchResultsGroup)
$form.controls.add($closeButton)

$MsgBoxNotify = [System.Windows.Forms.MessageBox]

#region events {
#$ConnectButton.Add_Click({Connect})
$SearchButton.Add_Click({Search})
$LogsButton.Add_Click({CreateLog})
$DeleteButton.Add_Click({DeleteEmails})
$closeButton.Add_Click({ closeForm })
#endregion events

#endregion GUI
$global:StartDateString = $null
$global:StartDate = $null
$global:EndDateString = $null
$global:EndDate = $null
$global:QuerySubject = $null
$global:QuerySender = $null
$global:ContentQuery = $null
$global:namestatus = $null



function radiobutton_CheckedChanged 
{
    
    $SearchLocationTxtbox.Enabled   = $True
    if($radiobutton1.Checked)
    {
        $EmailAddressTextBox.Enabled    = $True
        $SubjectTxtbox.Enabled          = $True
        $DaysTextBox.Enabled            = $True
    }
    elseif ($radiobutton2.Checked)
    {
        $EmailAddressTextBox.Enabled    = $False
        $SubjectTxtbox.Enabled          = $True
        $DaysTextBox.Enabled            = $True
    }
    elseif ($radiobutton3.Checked)
    {
        $EmailAddressTextBox.Enabled    = $True
        $SubjectTxtbox.Enabled          = $False
        $DaysTextBox.Enabled            = $True
    }
    else
    {
        $EmailAddressTextBox.Enabled    = $True
        $SubjectTxtbox.Enabled          = $True
        $DaysTextBox.Enabled            = $False
    }
}
    

#Search functions

#Calucalte dates range
function Date_Range()
{
    $global:StartDateString = (Get-Date).AddDays(-$DaysTextBox.text).ToShortDateString()
    $global:StartDate = $global:StartDateString | Get-Date -Format yyyy-MM-dd
    $global:EndDateString = (Get-Date).ToShortDateString()
    $global:EndDate = $global:EndDateString | Get-Date -Format yyyy-MM-dd
}

function Search ()
{
    if($radiobutton1.Checked)
    {
        Search_All
    }
    elseif($radiobutton2.Checked)
    {
        Search_Sub_Date
    }
    elseif($radiobutton3.Checked)
    {
        Search_Send_Date
    }
    else
    {
        Search_Sub_Send
    }
}

#Search by Subject and Sender and Date range
function Search_All()
{ 
    #Check that all the fileds are filled

    if ( $SubjectTxtbox.Text -and $EmailAddressTextBox.text -and $DaysTextBox.text -and $SearchLocationTxtbox.text)
    {
        Date_Range
        $global:QuerySubject = $SubjectTxtbox.Text
        $global:QuerySender = $EmailAddressTextBox.Text
        $global:ContentQuery = "(c:c)(date=$global:StartDate..$global:EndDate)(subjecttitle='$global:QuerySubject')(senderauthor=$global:QuerySender)"
        search_process
    }
    else
    {
        $ResultTxtbox.text = "`r`n" +$str036
    }
}

#Search by Subject and Date range
function Search_Sub_Date()
{
    #Check that all the fileds are filled
    if ( $SubjectTxtbox.Text -and $DaysTextBox.text -and $SearchLocationTxtbox.text)
    {
        Date_Range
        $global:QuerySubject = $SubjectTxtbox.Text
        $global:QuerySender = $EmailAddressTextBox.Text
        $global:ContentQuery = "(c:c)(date=$global:StartDate..$global:EndDate)(subjecttitle='$global:QuerySubject')"
        search_process
    }
    else
    {
        $ResultTxtbox.text  = "`r`n" +$str036
    }
}

#Search by Subject and Sender
function Search_Sub_Send()
{
    #Check that all the fileds are filled
    if ( $SubjectTxtbox.Text -and $EmailAddressTextBox.text -and $SearchLocationTxtbox.text)
    {
        $global:QuerySubject = $SubjectTxtbox.Text
        $global:QuerySender = $EmailAddressTextBox.Text
        $global:ContentQuery = "(c:c)(subjecttitle='$global:QuerySubject')(senderauthor=$global:QuerySender)"
        search_process
    }
    else
    {
        $ResultTxtbox.text = "`r`n"  +$str036
    }
}

function Search_Send_Date()
{
    #Check that all the fileds are filled
    if ( $EmailAddressTextBox.text -and $DaysTextBox.text -and $SearchLocationTxtbox.text)
    {
        Date_Range
        $global:QuerySubject = $SubjectTxtbox.Text
        $global:QuerySender = $EmailAddressTextBox.Text
        $global:ContentQuery = "(c:c)(date=$global:StartDate..$global:EndDate)(senderauthor=$global:QuerySender)"
        search_process
    }
    else
    {
        $ResultTxtbox.text = "`r`n"  +$str036
    }
}

#Running the search
function search_process ()
{
    New-ComplianceSearch -Name $SearchNameTextBox.text -ExchangeLocation $SearchLocationTxtbox.text -ContentMatchQuery $global:ContentQuery
    $SearchStatus.Visible = $true
    $SearchStatus.Text = $str030
    $SearchStatus.Forecolor = $str039
    Start-ComplianceSearch -Identity $SearchNameTextBox.text
    Get-ComplianceSearch -Identity $SearchNameTextBox.text
    # Show the search status
    do
        {
            $ComplianceSearchStatus = Get-ComplianceSearch -Identity $SearchNameTextBox.text
            Start-Sleep 2
            $SearchStatus.Text = "Searching"
            $SearchStatus.Forecolor = $str033
        }
    until ($ComplianceSearchStatus.status -match "Completed")
           $SearchStatus.Text        = $str040
           $SearchStatus.Forecolor   = $str038 
           $LogsButton.visible       = $True
           $SearchButton.visible     = $false
}


# Create the log
function CreateLog()
{
    $SearchStatus.Text = $str031
    $SearchStatus.Forecolor = $str039 
    $ComplianceSearchStatus = Get-ComplianceSearch -Identity $SearchNameTextBox.text
    $ItemsCount = Get-ComplianceSearch -Identity $SearchNameTextBox.text | Select-Object items -ExpandProperty items
    $SearchResults = $ComplianceSearchStatus.SuccessResults
    if ($ComplianceSearchStatus.items -le 0)
    {
        $SearchStatus.Text = "The Search didn't return any useful results"
        $SearchStatus.Forecolor = $str037
    }
    else
    {
        $mailboxes = @()
        $SearchResultsLines = $SearchResults -split '[\r\n]+'
        foreach ($SearchResultsLine in $SearchResultsLines)
        {
            # If the Search Results Line matches the regex, and $matches[2] (the value of "Item count: n") is greater than 0)
            if ($SearchResultsLine -match 'Location: (\S+),.+Item count: (\d+)' -and $ComplianceSearchStatus.items -gt 0)
            {
                # Add the Location: (email address) for that Search Results Line to the $mailboxes array
                $mailboxes += $matches[1]
            }
            $mailboxes | out-file ~\desktop"\AffectedMailboxes.csv"
        } 
        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.Workbooks.Open($LogFullPath)
        $objExcel.Visible = $true
        $SearchStatus.Text = " "+$ItemsCount+ " items have been found `r`n The log file can be found at `r`n" +$LogFullPath
        $SearchStatus.Forecolor = $str038 
        $LogsButton.visible       = $false
        $DeleteButton.visible     = $True
    }
}


#Delete the emails
function DeleteEmails ()
{
    $ClickResult = $MsgBoxNotify::Show('WARNING: THERE IS NO AUTOMATED WAY TO RESTORE THESE ITEMS IF YOU DELETE THEM!"','Confirm Deletion','OkCancel','Warning')
    $PurgeSuffix = "_purge"
    $PurgeName = $SearchNameTextBox.text + $PurgeSuffix
    if ($ClickResult -eq 'Ok')
    {
        New-ComplianceSearchAction -SearchName  $SearchNameTextBox.text -Purge -PurgeType HardDelete -Confirm:$false
        do
        {
            $PurgeProcess = Get-ComplianceSearchAction -Identity $PurgeName
            Start-Sleep 2
            $SearchStatus.Text = "Purging"
            $SearchStatus.Forecolor = $str039 
        }
        until ($PurgeProcess.Status -match "Completed")
        Start-Sleep 2
        $PurgeProcess.Results | out-file ~\desktop"\PurgeResults.csv"
        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.Workbooks.Open($PurgeFullPath)
        $objExcel.Visible       = $true
        $SearchStatus.Text      = "The purge has been completed `r`n Purge log can be found at `r`n" +$PurgeFullPath
        $SearchStatus.Forecolor = $str038 
        $SearchButton.visible   = $True
        $LogsButton.visible     = $false
        $DeleteButton.visible   = $false
    }
    else
    {
        return
    }
    Remove-ComplianceSearch -Identity $SearchNameTextBox.text -Confirm:$false
}

function closeForm(){$Form.close()}

[void]$Form.ShowDialog()