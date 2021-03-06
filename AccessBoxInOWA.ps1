param(
	[string]$name = ""
)

#Gets Credentials stored in a encripted file
function Connect(){
    $credentialsfile = "C:\temp\credfile.txt"

    if (Test-Path $credentialsfile){  
        $password = cat $credentialsfile | convertto-securestring
        $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "Admin@", $password
    }else{
        $creds = Get-Credential –credential Admin@
        $encp = $creds.password 
        $encp |ConvertFrom-SecureString |Set-Content $credentialsfile
    }    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session
}#Close Function Connect



#Opens Internet explorer and puts in password
#Once this function is started you need to leave you mouse still or it could have unexpected results
function OpenIE([string]$url){
    $wshell = New-Object -com WScript.Shell
    $wshell.Run("iexplore.exe $url")    
    Start-Sleep 5
    $wshell.sendkeys("AdminUPN@asdfasd.com")
    $wshell.sendkeys("{TAB}")
    $wshell.sendkeys("Password")
    Start-Sleep 2
    $wshell.sendkeys("{TAB}")
    $wshell.sendkeys("{ENTER}")
}#Close Function OpenIE



#For opening in chrome
function OpenChrome([string]$url){
    $wshell = New-Object -com WScript.Shell
    $wshell.Run("chrome.exe $url")
}#Close Function OpenChrome



#For checking rights
function hasRights([string]$name){
    foreach($user in (Get-MailboxPermission $name)){if($user.user -eq "PROD\admin52290-959343680"){return $True}}
    return $False
}#Close Function hasRights



function mailboxExists([string]$name)
{
    $($count = @(Get-Mailbox $name).count) 2>&1 | out-null
    if($count -eq 0){
        Write-host "ERROR: Mailbox `"$name`" Does Not Exist" -ForegroundColor red          
    }elseif($count -gt 1){
        Write-host "ERROR: More Than One Mailbox Was Found Using: $name `n`rDo Not Use Regular Expressions" -ForegroundColor red
    }else{
            return $True
    }
        
    return $False
}#Close Function mailboxExists

####
# Script Starts Here
####

#Checks if a session is open if it isn't it opens one
if((Get-PSSession).state -ne "Opened"){Connect}

#If no input was given in the parameter then it will ask for the address
if($name -eq ""){$name = Read-Host 'Email Address:'}

if((mailboxExists $name) -eq $True){
    #If the rights aren't already granted it grants them
    if((hasRights $name) -eq $False){$(Add-MailboxPermission -user admin@webmail.byui.edu -identity $name -AccessRights fullaccess) 2>&1 | out-null}

    $url = "https://outlook.office365.com/owa/$name"

    #Calls functions to open the browser
    #IE commented out because it was mainly a test
    $(openChrome $url) 2>&1 | out-null
    #openIE $url
}#Close mailboxExists If statement