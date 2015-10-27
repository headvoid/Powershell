param (
	[bool]$ProcessSubfolders = $true,
	[bool]$DeleteSourceFolder = $true,
    [bool]$Impersonate = $true,
	[string]$EWSManagedApiPath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll",
	[bool]$LogVerbose = $false
);

# Define our functions

Function ShowParams()
{
	Write-Host "MoveItems -Mailbox <string>";
	Write-Host "";
	Write-Host "Required:";
	Write-Host " -Mailbox : Mailbox SMTP email address";
	Write-Host "";
	Write-Host "Optional:";
	Write-Host " -ProcessSubfolders : If true, subfolders of the source folder will also be processed (default is false)";
	Write-Host " -DeleteSourceFolder : If true, the source folder will be deleted once items moved (so long as it is empty)";
	Write-Host " -Username : Username for the account being used to connect to EWS (if not specified, current user is assumed)";
	Write-Host " -Password : Password for the specified user (required if username specified)";
	Write-Host " -Domain : If specified, used for authentication (not required even if username specified)";
	Write-Host " -Impersonate : Set to $true to use impersonation.";
	Write-Host " -LogVerbose: Show detailed output";
	Write-Host "";
}


function DeleteItems()
{
	# Process all the items in the given source folder, and move them to the target
	
	if ($args -eq $null)
	{
		throw "No folders specified for MoveItems";
	}
	$SourceFolderObject = $args[0];
	$SourceFolderPath = $SourceFolderObject[0].DisplayName + '\' + $SourceFolderObject[1].DisplayName;
	
    if($SourceFolderPath -eq "\")
    {
        return
    }
	Write-Host "Deleting from", $SourceFolderPath -foregroundcolor White;
	
	# Set parameters - we will process in batches of 500 for the FindItems call
	$Offset=0;
	$PageSize=500;
	$MoreItems=$true;
	
    $View = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)
    $FindResults = $SourceFolderObject.FindItems($View)

	ForEach ($Item in $FindResults)
	{
		if ($LogVerbose) { Write-Host "Processing", $Item.Id.UniqueId -foregroundcolor Gray; }
		try
		{
			$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);
		}
		catch
		{
			Write-Host "Failed to delete item", $Item.Id.UniqueId -foregroundcolor Red
		}
	}

	# Now process any subfolders
	if ($SourceFolderObject.ChildFolderCount -gt 0)
	{
		$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000);
		$SourceFindFolderResults = $SourceFolderObject[1].FindFolders($FolderView);
        DeleteItems($SourceFindFolderResults, $SourceFolderPath);
	}
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath = $args[0];

    $RootFolder
	
	$Folder = $RootFolder;
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\';
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0);
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly;
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i]);
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View);
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null;
					Write-Host "Failed to find " $PathElements[$i];
					Write-Host "Requested folder path: " $FolderPath;
					break;
				}
				
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id)
			}
		}
	}
	
	$Folder;
}

##### String Converstions for folder enumeration

#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
}  

##### End String Converstions for folder enumeration

##### CREATE FOLDER SECTION 

[string]$info = "White"                # Color for informational messages
[string]$warning = "Yellow"            # Color for warning messages
[string]$error = "Red"                 # Color for error messages

function CreateFolder($MailboxName)
{
    $MailboxRoot = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
    $MailboxRoot.DisplayName = $FolderName

    #Call Save to actually create the folder
    $MailboxRoot.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

    Write-host "Folder Created for " $MailboxName -foregroundcolor  $warning
}

$FolderName = ""

##### Move items between folders

function RemoveFolder($folderItem)
{
    $SourceFolderObject = GetFolder($MailboxRoot, $folderItem)
 
    if ($SourceFolderObject)
    {
	    # We have the source folder, now check we can get the target folder
	    if ($LogVerbose) { Write-Host "Source folder located: " $SourceFolderObject.DisplayName; }
	    
 		    # Found target folder, now initiate move
		    if ($LogVerbose) { Write-Host "Target folder located: " $TargetFolderObject.DisplayName; }
		    DeleteItems($SourceFolderObject);
		
		    # If delete parameter is set, check if the source folder is now empty (and if so, delete it)
			$SourceFolderObject.Load();

            
				# Folder is empty, so can be safely deleted
				try
				{
					$SourceFolderObject[1].Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::SoftDelete);
					Write-Host $folderItem "successfully deleted" -foregroundcolor Green;
				}
				catch
				{
					Write-Host "Failed to delete " $folderItem -foregroundcolor Red;
				}
   }
}


##### List Items in Mailbox

function FolderConfiguration($MailboxName)
{
    #Define Extended properties  
    $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
    $folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$Mailbox)  
    #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
    $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
    #Deep Transval will ensure all folders in the search path are returned  
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
    $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
    $PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
    $PR_DELETED_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  
    $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
    #Add Properties to the  Property Set  
    $psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED);  
    $psPropertySet.Add($PR_Folder_Path);  
    $fvFolderView.PropertySet = $psPropertySet;  
    #The Search filter will exclude any Search Folders  
    $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
    $fiResult = $null  
    #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
    
    $myArray = @()

    do {  
        $fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
        foreach($ffFolder in $fiResult.Folders){  
            if(($ffFolder.displayName -eq "Inbox") -or ($ffFolder.displayName -eq "Junk EMail") -or ($ffFolder.displayName -eq "Junk E-Mail") -or ($ffFolder.displayName -eq "Contacts") -or ($ffFolder.displayName -eq "Junk E-Mail") -or ($ffFolder.displayName -eq "Deleted Items") -or ($ffFolder.displayName -eq "Sent Items") -or ($ffFolder.displayName -eq "RSS Feeds") -or ($ffFolder.displayName -eq "Notes") -or ($ffFolder.displayName -eq "Outbox") -or ($ffFolder.displayName -eq "Tasks") -or ($ffFolder.displayName -eq "Drafts") -or ($ffFolder.displayName -eq "Sync Issues") -or ($ffFolder.displayName -eq "Journal") -or ($ffFolder.displayName -eq "Calendar")-or ($ffFolder.displayName -eq "Conversation Action Settings") -or ($ffFolder.displayName -eq "Working Set")) {
                # Write-Host "Skipping Folder :" $ffFolder.displayName
            }
            else {
                $myArray += $ffFolder.displayName              
            }
        } 
        $fvFolderView.Offset += $fiResult.Folders.Count
    }while($fiResult.MoreAvailable -eq $true)

    foreach($item in $myArray){
        $FolderName = $item

        RemoveFolder($item)
    }
    
    $ContactsFolderid  = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$_)
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $findResults = $Service.FindItems($ContactsFolderid,$view)
    foreach($contact in $findResults)
    {
        $contact.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::HardDelete)
    }
    $CalendarFolderid  = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$_)
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $findResults = $Service.FindItems($CalendarFolderid,$view)
    foreach($appointment in $findResults)
    {
        $appointment.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::HardDelete)
    }
    $Repeat = $true

    while($Repeat)
    {
        $InboxFolderid  = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$_)
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $findResults = $Service.FindItems($InboxFolderid,$view)
        Write-Host "Remaining in Inbox : "$findResults.TotalCount
        foreach($inbox in $findResults)
        {
            $inbox.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::HardDelete)
        }
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $findResults = $Service.FindItems($InboxFolderid,$view)
        If($findResults.TotalCount -eq 0)
        {
            $Repeat=$false
        }

    }
    $Repeat = $true

    while($Repeat)
    {
        $SentFolderid  = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$_)
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $findResults = $Service.FindItems($SentFolderid,$view)
        foreach($sent in $findResults)
        {
            $sent.Delete([Microsoft.Exchange.Webservices.Data.DeleteMode]::HardDelete)
        }
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $findResults = $Service.FindItems($SentFolderid,$view)
        Write-Host "Remaining in Sent Items : "$findResults.TotalCount
        If($findResults.TotalCount -eq 0)
        {
            $Repeat=$false
        }
    }
}


##### END CREATE FOLDER SECTION 

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
                  
$dllfolderpath = "C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll"
Write-Host "Using" $dllfolderpath

[void][Reflection.Assembly]::LoadFile($dllfolderpath);
Add-Type -Path $dllfolderpath

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1;
  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion);

 $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials("justin.whelan@douglas.co.nz","F0r31gn12!");  
 $uri = [system.URI] "https://outlook.office365.com/EWS/Exchange.asmx"
 $service.Url = $uri;
 
#$users = Get-Content '..\Douglas Pharmaceuticals\O365Cleanup\users.txt'  ## <--- CHANGE THIS TO MATCH THE ACTUAL FILE
#$users = @("Liw@nhlab.co.nz","bob@abc.com")
$count = 0

foreach($person in $users) {     

    $Mailbox = $person
    $person
    if ($Impersonate)
    {
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox);
    }

    $goodtogo=$false
    
    if($service.ImpersonatedUserId)
    {
        try {
            $MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot);
            $InboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox);
            $ContactRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts);
            $SentItemsRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems);
            $SentItemsRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar);
            $goodtogo=$true
 
        }
        catch {
            write-host "problem connecting to" $Mailbox
            Add-Content '..\Douglas Pharmaceuticals\O365Cleanup\failed.txt' $Mailbox
            
        }
        if($goodtogo -eq $true)
        {
            Write-Host "Connected to" $Mailbox
            FolderConfiguration($Mailbox)
            Add-Content '..\Douglas Pharmaceuticals\O365Cleanup\cleaned.txt' $Mailbox

        }
    }
}
