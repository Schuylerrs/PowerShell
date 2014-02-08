##########################################################################
#
# This is a work in progress, it makes an MAPI connection to a server and
# removes emails with a specified MessageID from their inbox.
#
# This uses code found at http://poshcode.org/624
#
# Changes needed:
# - Make mailboxes be an input (they are hard coded in right now)
# - Grant rights to the mailbox before attempting to open the mailbox
#
##########################################################################



## Start code from http://poshcode.org/624

## EWS Managed API Connect Script
## Requires the EWS Managed API and Powershell V2.0 or greator  
  
## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version (Exchange2007 SP1)
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1  
  
## Create Exchange Service Object 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)   
$service.UseDefaultCredentials = $true  
  
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use      
 

 # An account must be used to 
$Subject = Read-Host "Message ID:"
$Names = @("testemail3@ourDomain.com","testemail2@ourDomain.com","testemail@ourDomain.com")
ForEach($name in $Names){
    $name
    $service.AutodiscoverUrl($name)  
    "Using CAS Server : " + $Service.url


    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$name)  
    $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

    $AqsString = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,$Subject)

    $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
    $fiItems = $null    
    do{    
        $fiItems = $Inbox.FindItems($AqsString,$ivItemView)    
        #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
        foreach($Item in $fiItems.Items){      
            # Delete the Message  
            $Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
            "Last Mail From : " + $Item.From.Name
            "Subject : " + $Item.Subject
            "Sent : " + $Item.DateTimeSent
            "ID : " + $Item.InternetMessageId
        }    
        $ivItemView.Offset += $fiItems.Items.Count    
    }while($fiItems.MoreAvailable -eq $true)    
}


#Playing with creating a folder

function CreateFolder()
{
    #Define Folder Name to Search for  
    $FolderName = "My New Folder123"  
    #Define Folder Veiw Really only want to return one object  
    $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)  
    #Define a Search folder that is going to do a search based on the DisplayName of the folder  
    $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$FolderName)  
    #Do the Search  
    $findFolderResults = $service.FindFolders($Inbox.id,$SfSearchFilter,$fvFolderView)  

    if ($findFolderResults.TotalCount -eq 0){  
        "Folder Doesn't Exist"
        ## Bind to the Inbox Sample  
  
        $NewFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)  
        $NewFolder.DisplayName = "My New Folder123"  
        $NewFolder.Save($Inbox.Id)    
    }  
    else{  
        "Folder Exist"  
    }
}  