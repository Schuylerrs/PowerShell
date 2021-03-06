##############################################################################
#
# After a phishing outbreak we had some user accounts comprimised
# I created this to mitigate the problem and monitor for infected users
# 
# It checks the queue to see if people have items stuch in the Queue because 
# the infected users were often sending to addresses that weren't valid
# 
# Mailboxes above a given threshold were disabled
# A message is sent to the admins when mailboxes come close to of pass the
# threshold
#
# This was scheduled to run every 5 minutes and the thresholds were arbitrarily
#
##############################################################################

$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://FQDN.Address/PowerShell/ -Credential $cred
Import-PSSession $s

$server = Get-ExchangeServer * | where{$_.ServerRole -eq "HubTransport"}
$messages = $server | %{Get-Message -Server $_.identity}
($messages = $messages | where{$_.fromaddress -like "*@ourDomain.com"} | where{(get-mailbox $_.fromaddress).ExchangeUserAccountControl -eq "None"}) 2>&1 | out-null
#Only top 10 senders are displayed in the email
$messages = $messages | group fromaddress | sort count -Descending | select count,name,group -first 10
    
#Output if formated in HTML to give a nice looking table
$text = "<h2>Top Caught In Queue</h2><br>"
$text += "<table border=`"1`"><tr><th>Count</th><th>Email Address</th></tr>"
$text += $messages | %{"<tr><td>$($_.Count)</td><td>$($_.name)</td></tr>"}
$text += "</table>"
$text += $messages | %{
    "<br><br>___________________________________________<br>" + $_.name + " Email Subjects<br>_________________________________________<br>";
    $_.group | sort subject -Unique | select subject -First 20 | %{"<br>" + $_.subject + "<br>"}
}
    
#Checks if people are sending large amounts of email (As a warning only)
$senders = $servers | %{get-messagetrackinglog -EventId "SEND" -server $_.name -Start (get-date).addhours(-1) -ResultSize unlimited} | where{($_.recipients -notlike "*ourDomain*")} | group sender | sort count -Descending | select count, name -First 10
$text += "<h2>Top Senders To External</h2><br>"
$text += "<table border=`"1`"><tr><th>Count</th><th>Email Address</th></tr>"
$text += $senders | %{"<tr><td>$($_.Count)</td><td>$($_.name)</td></tr>"}
$text += "</table>"
$text += $senders | %{
    "<br><br>___________________________________________<br>" + $_.name + " Email Subjects<br>_________________________________________<br>";
    $_.group | sort subject -Unique | select subject -First 20 | %{"<br>" + $_.subject + "<br>"}
}
    
    
#Sending email 
Send-MailMessage -To "emailteam@ourDomain.com" -Subject ("Email Queue as of " + [string](get-date)) -Body $text -SmtpServer "smtp.FQDN.com" -Credential $cred -BodyAsHtml
