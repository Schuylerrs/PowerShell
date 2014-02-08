####################################################################################
#
# This finds all people that recieved a specific email, grants rights to the account
# and removes the email from their inbox.
#
# This was mainly to learn how to use GUI objects
#
####################################################################################

function SeekAndDestroy
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Enter All Message Information"
    $objForm.Size = New-Object System.Drawing.Size(300,300) 
    $objForm.StartPosition = "CenterScreen"

    $x = "balbalbaba"

    $objForm.KeyPreview = $True

    $objMessageIDText = New-Object System.Windows.Forms.Label
    $objMessageIDText.Location = New-Object System.Drawing.Size(10,10) 
    $objMessageIDText.Size = New-Object System.Drawing.Size(380,20) 
    $objMessageIDText.Text = "Message ID:"
    $objForm.Controls.Add($objMessageIDText) 

    $objMessageIDBox = New-Object System.Windows.Forms.TextBox 
    $objMessageIDBox.Location = New-Object System.Drawing.Size(10,30) 
    $objMessageIDBox.Size = New-Object System.Drawing.Size(260,20) 
    $objForm.Controls.Add($objMessageIDBox)

    $objSubjectText = New-Object System.Windows.Forms.Label
    $objSubjectText.Location = New-Object System.Drawing.Size(10,60) 
    $objSubjectText.Size = New-Object System.Drawing.Size(380,20) 
    $objSubjectText.Text = "Message Subject:"
    $objForm.Controls.Add($objSubjectText) 

    $objSubjectBox = New-Object System.Windows.Forms.TextBox 
    $objSubjectBox.Location = New-Object System.Drawing.Size(10,80) 
    $objSubjectBox.Size = New-Object System.Drawing.Size(260,20) 
    $objForm.Controls.Add($objSubjectBox) 

    $objSenderText = New-Object System.Windows.Forms.Label
    $objSenderText.Location = New-Object System.Drawing.Size(10,110) 
    $objSenderText.Size = New-Object System.Drawing.Size(380,20) 
    $objSenderText.Text = "Message Sender:"
    $objForm.Controls.Add($objSenderText) 

    $objSenderBox = New-Object System.Windows.Forms.TextBox 
    $objSenderBox.Location = New-Object System.Drawing.Size(10,130) 
    $objSenderBox.Size = New-Object System.Drawing.Size(260,20) 
    $objForm.Controls.Add($objSenderBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(112,220)
    $OKButton.Size = New-Object System.Drawing.Size(75,25)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $objForm.Topmost = $True

    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$objForm.Close()}})

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    $MsgID = ""
    $MsgSubject = ""
    $MsgSender = ""

    $MsgID = $objMessageIDBox.Text
    $MsgSubject = $objSubjectBox.Text
    $MsgSender = $objSenderBox.Text

    "************************************************************************************************"
    "Searching for messages"
    "Sender: " + $MsgSender
    "Subject: " + $MsgSubject
    "ID: " + $MsgID
    "`nRecipients Found:"
   
    
    $servers = Get-ExchangeServer * | where{$_.ServerRole -eq "HubTransport"}
    #Builds the command to run based on input given
    $CMD = "`$Recipients = `$servers | %{get-messagetrackinglog -server `$_.name"
    
    If($MsgSubject -ne "") {$CMD += " -MessageSubject `$MsgSubject"}
    If($MsgID -ne "") {$CMD += " -MessageId `$MsgID"}
    If($MsgSender -ne "") {$CMD += " -Sender `$MsgSender"}

    $CMD += " -Start (get-date).AddDays(-5)} | select recipients, messagesubject"
    
    Invoke-Expression $CMD

    If($MsgSubject -eq "") {$MsgSubject = $Recipients[0].messagesubject}
    
    $Recipients = $Recipients  | %{$_.Recipients | %{$_}} | sort -Unique | where{@(get-mailbox $_ -ErrorAction "SilentlyContinue").count -eq 1}
    $Recipients
    $Recipients | %{
        Add-MailboxPermission -Identity $_ -User "admin" -AccessRights FullAccess
        Get-Mailbox $_ | Export-Mailbox –SubjectKeywords $MsgSubject –TargetMailbox "admin@ourDomain.com" –TargetFolder "test 2" -DeleteContent | Out-Null
    }
    "************************************************************************************************"
}


SeekAndDestroy