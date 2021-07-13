<#

***************   This script is provided AS-IS without any warranty to any damage that may occured. It will delete emails, if you are using it it's AT YOUR OWN RISK!  ***********


Version 2.0

As Microsoft changed the search way I re-write the code. The code is now using the Office 365 Security & Compliance


Version 1.1
Check if the Exchange Online PowerShell using multi-factor authentication module is installed

Version 1.0
Inital release


#>


$path = Resolve-Path ~\desktop
$file = '\AffectedMailboxes.csv'
$fullpath = Join-Path -path $path -ChildPath $file
$str001 = "Remove Malicious Emails Ver 2.0"
$str002 = "***********    Please make sure you are a member of the eDiscovery Manager role and the Organization Management groups    **********"
$str003 = "To check if you are a member of one of those groups go to https://protection.office.com and on the left pane click  Permissions"
$str004 = "I will search in your entire environment, but I will return only a maximum of 500 Mailboxes in the results."
$str005 = "The Purge Action will purge a maximum of 10 Items per Mailbox."
$str006 = "I will move items to the user's Recoverable Items folder, and they will remain there based on the Retention Period that is configured for the mailbox (searching will return also those emails)."
$str007 = "Please Connect to the Office 365 Security & Compliance"
$str008 = "Connect with MFA Account"
$str009 = "Connect without MFA Account"
$str010 = "Connection Status:"
$str011 = "Not Connected to the Security and Compliance Center"
$str012 = "Type a Name for this search (less than 198 Characters)"
$str013 = "Sender Email Address"
$str014 = "Message Subject"
$str015 = "Days To Search"
$str016 = "Where to search?"
$str017 = "Type Email Address of MailBox or Group you would like to search within (to search in all MB's type all)"
$str018 = "What type of search you going perform?"
$str019 = "By Subject, Sender Address and Date Range"
$str020 = "By Subject and Date Range"
$str021 = "By Subject and Sender Address"
$str022 = "Get a list of the affected mailboxes"
$str023 = "Search Status:"
$str024 = "Delete the emails"
$str025= "The search didn't created yet"
$str026 = "Results"
$str027 = "Close"
$str028 = "Created the search... please wait for it to start"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1000,850'
$Form.text                       = $str001
$Form.TopMost                    = $false

$MFA_AccountButton               = New-Object system.Windows.Forms.Button
$MFA_AccountButton.text          = $str008
$MFA_AccountButton.width         = 200
$MFA_AccountButton.height        = 30
$MFA_AccountButton.location      = New-Object System.Drawing.Point(270,60)
$MFA_AccountButton.Font          = 'Microsoft Sans Serif,10'

$NonMFA_AccountButton            = New-Object system.Windows.Forms.Button
$NonMFA_AccountButton.text       = $str009
$NonMFA_AccountButton.width      = 200
$NonMFA_AccountButton.height     = 30
$NonMFA_AccountButton.location   = New-Object System.Drawing.Point(540,60)
$NonMFA_AccountButton.Font       = 'Microsoft Sans Serif,10'

$Status                          = New-Object system.Windows.Forms.Label
$Status.text                     = $str010
$Status.AutoSize                 = $true
$Status.width                    = 25
$Status.height                   = 10
$Status.location                 = New-Object System.Drawing.Point(280,110)
$Status.Font                     = 'Microsoft Sans Serif,10'

$StatusUpdate                    = New-Object system.Windows.Forms.Label
$StatusUpdate.ForeColor          = 'Red'
$StatusUpdate.Text               = $str011
$StatusUpdate.AutoSize           = $true
$StatusUpdate.width              = 25
$StatusUpdate.height             = 10
$StatusUpdate.location           = New-Object System.Drawing.Point(400,110)
$StatusUpdate.Font               = 'Microsoft Sans Serif,10'

$SearchName                      = New-Object system.Windows.Forms.Label
$SearchName.text                 = $str012
$SearchName.AutoSize             = $true
$SearchName.width                = 25
$SearchName.height               = 10
$SearchName.location             = New-Object System.Drawing.Point(20,147)
$SearchName.Font                 = 'Microsoft Sans Serif,10'

$SearchNameTextBox               = New-Object system.Windows.Forms.TextBox
$SearchNameTextBox.multiline     = $false
$SearchNameTextBox.width         = 200
$SearchNameTextBox.height        = 20
$SearchNameTextBox.location      = New-Object System.Drawing.Point(365,145)
$SearchNameTextBox.Font          = 'Microsoft Sans Serif,10'

$EmailAddress                    = New-Object system.Windows.Forms.Label
$EmailAddress.text               = $str013
$EmailAddress.AutoSize           = $true
$EmailAddress.width              = 25
$EmailAddress.height             = 10
$EmailAddress.location           = New-Object System.Drawing.Point(20,189)
$EmailAddress.Font               = 'Microsoft Sans Serif,10'

$EmailAddressTextBox             = New-Object system.Windows.Forms.TextBox
$EmailAddressTextBox.multiline   = $false
$EmailAddressTextBox.width       = 645
$EmailAddressTextBox.height      = 20
$EmailAddressTextBox.location    = New-Object System.Drawing.Point(200,187)
$EmailAddressTextBox.Font        = 'Microsoft Sans Serif,10'

$Subject                         = New-Object system.Windows.Forms.Label
$Subject.text                    = $str014
$Subject.AutoSize                = $true
$Subject.width                   = 25
$Subject.height                  = 10
$Subject.location                = New-Object System.Drawing.Point(20,233)
$Subject.Font                    = 'Microsoft Sans Serif,10'

$SubjectTxtbox                   = New-Object system.Windows.Forms.TextBox
$SubjectTxtbox.multiline         = $false
$SubjectTxtbox.width             = 645
$SubjectTxtbox.height            = 20
$SubjectTxtbox.location          = New-Object System.Drawing.Point(200,231)
$SubjectTxtbox.Font              = 'Microsoft Sans Serif,10'

$Days                            = New-Object system.Windows.Forms.Label
$Days.text                       = $str015
$Days.AutoSize                   = $true
$Days.width                      = 25
$Days.height                     = 10
$Days.location                   = New-Object System.Drawing.Point(20,277)
$Days.Font                       = 'Microsoft Sans Serif,10'

$DaysTextBox                     = New-Object system.Windows.Forms.TextBox
$DaysTextBox.multiline           = $false
$DaysTextBox.width               = 60
$DaysTextBox.height              = 20
$DaysTextBox.location            = New-Object System.Drawing.Point(200,275)
$DaysTextBox.Font                = 'Microsoft Sans Serif,10'

$Where_Search                    = New-Object system.Windows.Forms.Label
$Where_Search.text               = $str016
$Where_Search.AutoSize           = $true
$Where_Search.width              = 25
$Where_Search.height             = 10
$Where_Search.location           = New-Object System.Drawing.Point(20,321)
$Where_Search.Font               = 'Microsoft Sans Serif,10'

$Where_Search2                   = New-Object system.Windows.Forms.Label
$Where_Search2.text              = $str017
$Where_Search2.AutoSize          = $true
$Where_Search2.width             = 25
$Where_Search2.height            = 10
$Where_Search2.location          = New-Object System.Drawing.Point(20,346)
$Where_Search2.Font              = 'Microsoft Sans Serif,10'

$Where_Search_Txtbox             = New-Object system.Windows.Forms.TextBox
$Where_Search_Txtbox.multiline   = $false
$Where_Search_Txtbox.width       = 200
$Where_Search_Txtbox.height      = 20
$Where_Search_Txtbox.location    = New-Object System.Drawing.Point(645,344)
$Where_Search_Txtbox.Font        = 'Microsoft Sans Serif,10'

$SerachType                      = New-Object system.Windows.Forms.Label
$SerachType.text                 = $str018
$SerachType.AutoSize             = $true
$SerachType.width                = 25
$SerachType.height               = 10
$SerachType.location             = New-Object System.Drawing.Point(360,400)
$SerachType.Font                 = 'Microsoft Sans Serif,10'

$Search_All                      = New-Object system.Windows.Forms.Button
$Search_All.text                 = $str019
$Search_All.width                = 300
$Search_All.height               = 60
$Search_All.location             = New-Object System.Drawing.Point(50,430)
$Search_All.Font                 = 'Microsoft Sans Serif,10'

$Search_Sub_Date                 = New-Object system.Windows.Forms.Button
$Search_Sub_Date.text            = $str020
$Search_Sub_Date.width           = 222
$Search_Sub_Date.height          = 60
$Search_Sub_Date.location        = New-Object System.Drawing.Point(375,430)
$Search_Sub_Date.Font            = 'Microsoft Sans Serif,10'

$Search_Sub_Send                 = New-Object system.Windows.Forms.Button
$Search_Sub_Send.text            = $str021
$Search_Sub_Send.width           = 280
$Search_Sub_Send.height          = 60
$Search_Sub_Send.location        = New-Object System.Drawing.Point(620,430)
$Search_Sub_Send.Font            = 'Microsoft Sans Serif,10'

$LogsButton                      = New-Object system.Windows.Forms.Button
$LogsButton.text                 = $str022
$LogsButton.width                = 300
$LogsButton.height               = 60
$LogsButton.location             = New-Object System.Drawing.Point(50,500)
$LogsButton.Font                 = 'Microsoft Sans Serif,10'

$SearchStatus                    = New-Object system.Windows.Forms.Label
$SearchStatus.text               = $str023
$SearchStatus.AutoSize           = $true
$SearchStatus.width              = 25
$SearchStatus.height             = 10
$SearchStatus.location           = New-Object System.Drawing.Point(350,580)
$SearchStatus.Font               = 'Microsoft Sans Serif,10'
$SearchStatus.Visible            = $false

$DeleteButton                    = New-Object system.Windows.Forms.Button
$DeleteButton.text               = $str024
$DeleteButton.width              = 223
$DeleteButton.height             = 60
$DeleteButton.location           = New-Object System.Drawing.Point(375,500)
$DeleteButton.Font               = 'Microsoft Sans Serif,10'

$SearchStatusResults             = New-Object system.Windows.Forms.Label
$SearchStatusResults.ForeColor   = 'Red'
$SearchStatusResults.Text        = $str025
$SearchStatusResults.AutoSize    = $true
$SearchStatusResults.width       = 25
$SearchStatusResults.height      = 10
$SearchStatusResults.location    = New-Object System.Drawing.Point(480,580)
$SearchStatusResults.Font        = 'Microsoft Sans Serif,10'
$SearchStatusResults.Visible     = $false

$Message                         = New-Object system.Windows.Forms.Label
$Message.text                    = $str026
$Message.AutoSize                = $true
$Message.width                   = 25
$Message.height                  = 10
$Message.location                = New-Object System.Drawing.Point(20,681)
$Message.Font                    = 'Microsoft Sans Serif,10'

$result                          = New-Object system.Windows.Forms.TextBox
$result.multiline                = $true
$result.width                    = 800
$result.height                   = 190
$result.location                 = New-Object System.Drawing.Point(100,607)
$result.BackColor                = [System.Drawing.Color]::FromArgb(245,245,220) #, 192)
$result.text                     = "                                                                                  Welcome to "+ $str001 + "`r`n `r`n" +$str002 + "`r`n" +$str003 + "`r`n`r`n" +$str004 + "." +$str005 + "`r`n`r`n" +$str006 + "`r`n"
$result.Font                     = 'Microsoft Sans Serif,10'

$closeButton                     = New-Object system.Windows.Forms.Button
$closeButton.text                = $str027
$closeButton.width               = 102
$closeButton.height              = 30
$closeButton.location            = New-Object System.Drawing.Point(800,800)
$closeButton.Font                = 'Microsoft Sans Serif,10'

$MsgBoxError = [System.Windows.Forms.MessageBox]
$MsgBoxNotify = [System.Windows.Forms.MessageBox]

$Form.controls.AddRange(@($MFA_AccountButton,$NonMFA_AccountButton,$SearchName,$SearchNameTextbox,$EmailAddress,$EmailAddressTextBox,$Subject,$SubjectTxtbox,$Days,$DaysTextBox,$Where_Search,$Where_Search2,$Where_Search_Txtbox,$Message,$SerachType,$Status,$StatusUpdate,$Search_All,$Search_Sub_Date,$Search_Sub_Send,$LogsButton,$SearchStatus,$SearchStatusResults,$DeleteButton,$result,$closeButton))

#region events {
$MFA_AccountButton.Add_Click({ Connect_MFA_Account })
$NonMFA_AccountButton.Add_Click({ConnectNonMFA_Account})
$Search_All.Add_Click({ Search_All})
$Search_Sub_Date.Add_Click({Search_Sub_Date})
$Search_Sub_Send.Add_Click({Search_Sub_Send})
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


#Calucalte dates range
function Date_Range()
{
    $global:StartDateString = (Get-Date).AddDays(-$DaysTextBox.text).ToShortDateString()
    $global:StartDate = $global:StartDateString | Get-Date -Format yyyy-MM-dd
    $global:EndDateString = (Get-Date).ToShortDateString()
    $global:EndDate = $global:EndDateString | Get-Date -Format yyyy-MM-dd
}

#Search by Subject and Sender and Date range
function Search_All()
{ 
        Try
        {
            #Check that all the fileds are filled
                if ( $SubjectTxtbox.Text -and $EmailAddressTextBox.text -and $DaysTextBox.text -and $Where_Search_Txtbox.text)
                    {
                        Date_Range
                        $global:QuerySubject = $SubjectTxtbox.Text
                        $global:QuerySender = $EmailAddressTextBox.Text
                        $global:ContentQuery = "(c:c)(date=$global:StartDate..$global:EndDate)(subjecttitle='$global:QuerySubject')(senderauthor=$global:QuerySender)"
                        search_process
                    }
                else
                    {
                        $result.text = "`r`nSome fileds are missing data"
                    }
        }
        Catch [System.SystemException]
        {
                $MsgBoxError::Show($str007, $str001, "OK", "Error")
        }
}


#Search by Subject and Date range
function Search_Sub_Date()
{
    Try
    {
        #Check that all the fileds are filled
            if ( $SubjectTxtbox.Text -and $DaysTextBox.text -and $Where_Search_Txtbox.text)
                {
                    Date_Range
                    $global:QuerySubject = $SubjectTxtbox.Text
                    $global:QuerySender = $EmailAddressTextBox.Text
                    $global:ContentQuery = "(c:c)(date=$global:StartDate..$global:EndDate)(subjecttitle='$global:QuerySubject')(senderauthor=$global:QuerySender)"
                    search_process
                }
            else
                {
                    $result.text = "`r`nSome fileds are missing data"
                }
    }
    Catch [System.SystemException]
    {
        $MsgBoxError::Show($str007, $str001, "OK", "Error")
    }
}

#Search by Subject and Sender
function Search_Sub_Send()
{
    { 
        Try
        {
            #Check that all the fileds are filled
                if ( $SubjectTxtbox.Text -and $EmailAddressTextBox.text -and $Where_Search_Txtbox.text)
                    {
                        $global:QuerySubject = $SubjectTxtbox.Text
                        $global:QuerySender = $EmailAddressTextBox.Text
                        $global:ContentQuery = "(c:c)(subjecttitle='$global:QuerySubject')(senderauthor=$global:QuerySender)"
                        search_process
                    }
                else
                    {
                        $result.text = "`r`nSome fileds are missing data"
                    }
        }
        Catch [System.SystemException]
        {
            $MsgBoxError::Show($str007, $str001, "OK", "Error")
        }
    }
}

#Running the search
function search_process ()
{
    New-ComplianceSearch -Name $SearchNameTextBox.text -ExchangeLocation $Where_Search_Txtbox.text -ContentMatchQuery $global:ContentQuery
    $SearchStatus.Visible = $true
    $SearchStatusResults.Visible = $true
    $SearchStatusResults.Text = $str028
    $SearchStatusResults.Forecolor = 'Red'
    Start-ComplianceSearch -Identity $SearchNameTextBox.text
    Get-ComplianceSearch -Identity $SearchNameTextBox.text
    # Show the search status
    do
        {
            $ComplianceSearchStatus = Get-ComplianceSearch -Identity $SearchNameTextBox.text
            Start-Sleep 2
            $SearchStatusResults.Text = "Searching"
            $SearchStatusResults.Forecolor = 'Orange'
        }
    until ($ComplianceSearchStatus.status -match "Completed")
           $SearchStatusResults.Text = "Search has been completed"
           $SearchStatusResults.Forecolor = 'Green'
}

# Create the log
function CreateLog()
    {
        $ComplianceSearchStatus = Get-ComplianceSearch -Identity $SearchNameTextBox.text
        $SearchResults = $ComplianceSearchStatus.SuccessResults
        if ($ComplianceSearchStatus.items -le 0)
            {
                $result.text = "`r`n The Search didn't return any useful results "
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
                    $objExcel.Workbooks.Open($fullpath)
                    $objExcel.Visible = $true
                    $result.text = "`r`nThe log file can be found at " +$fullpath
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
                $SearchStatus.Visible = $true
                $SearchStatus.text = "Purging Status:"
                $SearchStatusResults.Text = "Purging"
            }

            until ($PurgeProcess.Status -match "Completed")
                
                $SearchStatusResults.Text = $PurgeProcess 
                Start-Sleep 2
                $SearchStatusResults.Text = $PurgeProcess.Status
                $result.text = "`r`n`r`n`r`n`r`n" +$PurgeProcess.Results
        }

        else
        {
            return
        }
        Remove-ComplianceSearch -Identity $SearchNameTextBox.text -Confirm:$false
    }


Function TestConnection()
{
    Try
    {
        Get-ComplianceSecurityFilter -ErrorAction Stop
        $StatusUpdate.ForeColor = 'Green'
        $StatusUpdate.Text = "Connected to the Security and Compliance Center"
    }
    Catch [System.SystemException]
    {
        $StatusUpdate.ForeColor = 'Red'
        $StatusUpdate.Text = "Not Connected to the Security and Compliance Center"
    }
}

function Get-ExchangeOnlineModule {
    [CmdletBinding()]  
    Param(
        $ApplicationName = "Microsoft Exchange Online Powershell Module"
    )
        $InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
        return $InstalledApplicationNotMSI | Where-Object { $_.displayname -match $ApplicationName } | Select-Object -First 1
    }

Function Test-ExchangeOnlineModule {
    [CmdletBinding()] 
    Param(
        $ApplicationName = "Microsoft Exchange Online Powershell Module"
    )
        return ( $null -ne (Get-ExchangeOnlineModule -ApplicationName $ApplicationName)) 
    }

 function MFA () 
    {
        if ((Test-ExchangeOnlineModule -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false) 
         {
            $MsgBoxError::Show('In order to connect with MFA account you need to install Exchange Online PowerShell multi-factor authentication',$str001,'Ok','Error')
        }
         else 
            {
                 try
                     {
                        $StatusUpdate.Text = "Connecting"
                        $StatusUpdate.Forecolor = 'Red'
                        Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)
                        $MFCCPSSession = New-ExoPSSession -ConnectionUri 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId' -UserPrincipalName $UPN
                        Import-PSSession $MFCCPSSession -AllowClobber
                    
                    }
                 catch
                    {
                        $MsgBoxError::Show("Wrong creds or no creds entered ...", $str001, "OK", "Error")
                    }
            }       
    }

function Non_MFA ()
    {
        try 
        {
            $StatusUpdate.Text = "Connecting"
            $StatusUpdate.Forecolor = 'Red'
            $UserCredential = Get-Credential -ErrorAction Continue
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
            Import-PSSession $Session -DisableNameChecking | Out-Null
        }
        catch
        {
            $MsgBoxError::Show("Wrong creds or no creds entered ...", $str001, "OK", "Error")
        }
    }

function Connect_MFA_Account ()
{
    MFA
    TestConnection
}

function ConnectNonMFA_Account ()
{
    Non_MFA
    TestConnection
}

function closeForm(){$Form.close()}

[void]$Form.ShowDialog()