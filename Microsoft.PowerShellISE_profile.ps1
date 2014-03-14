#############################################################
# Custom Functions
# These get defined every time powershell ISE opens
#############################################################

#############################################################
# Helper Functions
# Common tasks that are done often, used to add readablity 
#############################################################
# Checks if a box exists
function o365BoxExists($box)
{
    if(@(Get-Mailbox $box).count -eq 1){return $true}else{return $false}
}

# Tests if a remote session is open with a given computer
function isConnectedTo($connection)
{
    $session = Get-PSSession 
    if(($session.ComputerName -like $connection) -and ($session.state -eq "Opened")){return $true}else{return $false}
}

# Disconnects all remote sessions and exits the windows azure module
function disconnectAll($test)
{
    Exit-PSSession
	Get-PSSession | Remove-PSSession
    if((get-module).name -eq "MSOnline"){Remove-Module MsOnline}
}

# Not functioning
# Uses a encrypted file to store credential information
# Return comand not working as expected
function get-Cred($path, $account)
{
    if ((Test-Path $path) -eq $true)
    {
        Write-host "Path exists"  
        $password = cat $path | convertto-securestring
        $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password
    }else{
        Write-host "Path does not exists"
        $creds = Get-Credential –credential $account
        $encp = $creds.password 
        $encp | ConvertFrom-SecureString | Set-Content $path
    }
    Write-host "End get-cred"
    return $cred
}

#############################################################
# Main Connection Functions
# For connecting to various servers
#############################################################
# Connects to Office365
function ConnectO365($credentialsfile)
{ 
    if((isConnectedTo("*outlook.com*")) -eq $false)
    {
        disconnectAll($null)
        $credentialsfile = "C:\temp\credfile.txt"
        
        # Get-Cred function written in-line
        if (Test-Path $credentialsfile)
        {  
            $password = cat $credentialsfile | convertto-securestring
            $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "username@domain.com", $password
        }
        else
        {
            $creds = Get-Credential –credential username@domain.com
            $encp = $creds.password 
            $encp |ConvertFrom-SecureString | Set-Content $credentialsfile
        }    
        
        @($Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $cred -Authentication Basic -AllowRedirection) 2>&1 | out-null
        Import-PSSession $Session

        # Check to be sure it connected correctly
        if((isConnectedTo("dm2pr01psh.outlook.com")) -eq $false)
        {
            Write-Error "ERROR: Script failed to connect"
        }
    }
    else
    {
        Write-Host "Already Connected to O365"
    }
}

# Connects to our Exchange 2010 hybrid server
function ConnectExHybrid($test)
{
    # Check if connected already
    if((isConnectedTo("exhybrid.byui.edu")) -eq $false)
    { 
        disconnectAll($null)
        $s = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri http://Exhybrid.domain.com/PowerShell/ `
        -Authentication Kerberos
        Import-PSSession $s
    }
    else
    {
        Write-Host "Already Connected to Exchange Hybrid"
    }
}

# Opens the Azure module for running commands against the Office365 AD
function ConnectAzure($test)
{
    #Check if the module is already imported
    if((get-module).name -ne "MSOnline")
    {
        disconnectAll($null) 
        import-Module MsOnline
        $path = "C:\Users\dkadmin\Documents\blueCred.txt"

        # Get-Cred function written in-line
        if (Test-Path $credentialsfile)
        {  
            $password = cat $credentialsfile | convertto-securestring
            $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist "username@domain.com", $password
        }
        else
        {
            $creds = Get-Credential –credential username@domain.com
            $encp = $creds.password 
            $encp |ConvertFrom-SecureString | Set-Content $credentialsfile
        }    

        connect-msolservice -Credential $cred
    }
}

#############################################################
# Student Troubleshooting Functions
# For connecting to various servers
#############################################################
# Opens a student account in chrome (must be logged into your email first)
function OpenChrome($name)
{
    # Make sure that it is connected to office365
    if((isConnectedTo("bn1pr01psh.outlook.com")) -ne $true)
    {
        ConnectO365("C:\temp\credfile.txt")
    }

    $name = Read-Host 'Email Address:'
    
    if((o365BoxExists($name)) -eq $true)
    {
        Write-host "Accessing: $name"
        Get-MailboxStatistics $name
        Add-MailboxPermission -user username@domain.com -identity $name -AccessRights fullaccess
        Add-RecipientPermission $name -AccessRights SendAs -Trustee username@domain.com -Confirm:$false
        $url = "https://outlook.office365.com/owa/$name"
        
        # Creates a shell object (like cmd.exe)
        $wshell = New-Object -com WScript.Shell
        $wshell.Run("chrome.exe $url")
    }
    else
    {
        Write-Host "Mailbox `"$name`" does not exist!" -ForegroundColor red
    }
}

# Opens an employee account in chrome
function OpenEmployee($name)
{
    # 
    if((isConnectedTo("exhybrid.domain.com")) -eq $false)
    {
        ConnectExHybrid
    }

    $name = Read-Host 'Email Address:'
    
    try
    {   
        (get-mailbox $name -ErrorVariable stop) 2>&1 | out-null
        (Add-MailboxPermission -user username@domain.com -identity $name -AccessRights FullAccess -WarningAction silentlycontinue) 2>&1 | out-null
        (Add-MailboxPermission -user username@domain.com -identity $name -AccessRights SendAs -WarningAction silentlycontinue) 2>&1 | out-null
        Write-host "Accessing: $name" 
        Get-MailboxStatistics $name
        $url = "https://owa.byui.edu/owa/$name"
        $wshell = New-Object -com WScript.Shell
        $wshell.Run("chrome.exe $url")
    }
    catch
    {
        Write-Host "Mailbox `"$name`" does not exist!" -ForegroundColor red
    }
}
#############################################################
# Custom Add-ons Menu Items
# The create custom shortcuts
#############################################################
# Creates a sub menu
$connectMenu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Connect To...",$null,$null)

# Creating menu items
$connectMenu.SubMenus.Add(
  "Connect to Office 365",
    {
	    (ConnectO365("C:\temp\credfile.txt")) 2>&1 | out-null
    },
  "Control+Alt+O"
)

$connectMenu.SubMenus.Add(
  "Connect to ExHybrid",
    {
	    $(ConnectExHybrid("test")) 2>&1 | out-null
    },
  "Control+Alt+X"
)

$connectMenu.SubMenus.Add(
  "Connect to Azure",
  {
       $(ConnectAzure("test")) 2>&1 | out-null
  },
  "Control+Alt+Z"
)

$connectMenu.SubMenus.Add(
    "Check Connection",
    {
        Get-PSSession | ft ComputerName, State, Availability -AutoSize
    },
    "Control+Alt+C"
)

$connectMenu.SubMenus.Add(
  "Disconnect All",
    {
	    disconnectAll
    },
  "Control+Alt+D"
)

$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
   "Open Student Email (Chrome)",
    {
        $(OpenChrome($null)) 2>&1 | out-null
    },
    "Control+Alt+S"
)

$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add(
   "Get Stats",
    {
        $name = Read-Host 'Email Address:'
        Get-MailboxStatistics $name | ft
    },
    "Control+Alt+T"
)