$test = Get-Module Microsoft.Exchange.WebServices

if(!$test) {
    $dllInstallPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.2" -ErrorAction SilentlyContinue
    if(!$dllInstallPath) {
        $dllInstallPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -ErrorAction SilentlyContinue
    }
    if($dllInstallPath -ne $null) {
        $ewsDLLPath = $dllInstallPath.'Install Directory' + ".\Microsoft.Exchange.WebServices.dll"
        Import-Module $ewsDLLPath
    }
    else {
        Write-Error "Please install the Exchange Web Services API. Please download them here: https://www.microsoft.com/en-us/download/details.aspx?id=42951"
    }
}
Update-FormatData -PrependPath $PSScriptRoot\output.format.ps1xml
Function Connect-EXWebService {
    <#
    .SYNOPSIS
    Connects to Exchange Web Services. 
    
    .DESCRIPTION
    This cmdlet is used to connect to Exchange Web Services. By default it sets the ExchangeVersion and EWSUrl to connect to Office 365, however you can specify the arguments to an on premise Exchange System.

    This cmdlet is required to run any other EWS cmdlets such as Get-MailItem or Get-EXDelegate. It stores the connection as a global variable, so it can be accessed by all other cmdlets after running once.

    .PARAMETER Credential
    This parameter looks for a credential type object. It supposes a Get-Credential or any form of credential caching such as Get-SavedCreds in the CredManager module.

    .PARAMETER ExchangeVersion
    This parameter sets the Exchange version that EWS will use to connect. It primarily decides what features are available to you when programming. It defaults to Exchange2013_SP1, and some module features may not be available when run on 2010

    .PARAMETER EWSUrl
    This parameter sets the EWS Url, it must be formatted like "https://mail.domain.com/EWS/Exchange.asmx" It defaults to Office 365's EWS Url which is "https://outlook.office365.com/EWS/Exchange.asmx"

    .PARAMETER Force
    This parameter will allow you to recreate the connection object even if one exists. It is primarily used if unspecified errors start to occur.

    .EXAMPLE
    Connect-EXWebService -Credential (Get-Credential)

    .EXAMPLE
    Connect-ExWebService -Credential (Get-Credential) -ExchangeVersion "Exchange2010" -EWSUrl "https://mail.contoso.com/EWS/Exchange.asmx"

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [CmdletBinding()]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$true,Position=1)]
		[System.Management.Automation.PSCredential]$Credential,
		[Parameter(Mandatory=$false,Position=2)]
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]$ExchangeVersion="Exchange2013_SP1",
        [Parameter(Mandatory=$false)]
        [string]$EWSUrl="https://outlook.office365.com/EWS/Exchange.asmx",
		[Parameter(Mandatory=$false)]
		[switch]$Force
	)
	Process {
		#Try to get exchange service object from global scope
		$existingExSvcVar = (Get-Variable -Name exService -Scope Global -ErrorAction:SilentlyContinue) -ne $null
		
		#Establish the connection to Exchange Web Service
		if ((-not $existingExSvcVar) -or $Force) {
            $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    		 [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
			
			#Set network credential
			$userName = $Credential.UserName
            $password = $Credential.GetNetworkCredential().Password
			$exService.Credentials = New-Object System.Net.NetworkCredential($userName,$password)
			try {
                #Set the URL by using Autodiscover
                $exService.Url = $EWSUrl
				Set-Variable -Name exService -Value $exService -Scope Global -Force
			    }
			catch {
				$PSCmdlet.ThrowTerminatingError($_)
			    }
		    } 
        else {
            Write-Error "Exchange Connection already established!"
	        }
	    }
    }

Function Get-MailItem {
    <#
    .SYNOPSIS
    Searches a mailbox for Mail Items based on a number of criteria. 
    
    .DESCRIPTION
    This cmdlet is used to search for mail items in a given mailbox. It can search for emails, calendar entries (Both meetings and appointments), and contacts.

    It can search by Subject, To Address, From Address, and within a date range of time received (StartDate and EndDate). If you do not specify anything, it will return all messages in the specified folder.

    This cmdlet can be piped to either the Export-MailItem or Remove-MailItem to save or delete a located message respectively.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER Identity
    This parameter is used to specify the mailbox that will be searched. It allows the same types as Get-Mailbox (AKA alias, email address, and display name). This parameter is mandatory.

    .PARAMETER Folder
    This parameter specifies the folder to search. It accepts Inbox, Calendar, Contact, MsgFolderRoot, DeletedItems, and Root. MsgFolderRoot will search all folders in the mailbox that are a Note Folder Class (Folders containing mail messages), and Root will search all folders in the entire inbox. Root will search ALL folders, including attempting some system folders which will error. This parameter is mandatory

    .PARAMETER SubFolder
    This parameter allows you to specify a subfolder within the well known folder specified in the Folder parameter.

    .PARAMETER Subject
    This parameter allows you to specify a search string to search the subject of emails by. It is a substring search, so it will automatically assume there can be more text before and after the specified string without adding an asterisk on either side.

    .PARAMETER ToAddress
    This parameter lets you specify to whom the email was sent. This can be useful for looking for messages sent to a specific DL or sent to one person but forwarded on to another.

    .PARAMETER FromAddress
    This parameter lets you specify the sender of the message. 

    .PARAMETER StartDate
    This parameter allows you to specify a start datetime to begin looking. It can take standard datetime formats like "3/14/18" or "3-14-2018 3:34PM." It also accepts (Get-Date).AddDays(-2) or similar.

    .PARAMETER EndDate
    This parameter allows you to specify an end datetime to begin looking. It can take standard datetime formats like "3/14/18" or "3-14-2018 3:34PM." It also accepts (Get-Date).AddDays(-2) or similar.

    .PARAMETER ResultSize
    This parameter lets you set the result size of the number of emails you'd like to receive. It defaults to 1000, and will take any integer and "Unlimted."

    .EXAMPLE
    Get-MailItem -Identity jsmith -Folder Inbox -Subject "This is totally a virus!"

    .EXAMPLE
    Get-MailItem -Identity mmuhammad -Folder Calendar -Subject "Meeting that shouldn't exist!" -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date).AddDays(-13)

    .EXAMPLE
    Get-MailItem -Identity nholmes -Folder MsgFolderRoot -StartDate (Get-Date).AddDays(-1) -FromAddress "allCompany@contoso.com" | Remove-MailItem -DeleteType HardDelete -Confirm:$false 

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding(DefaultParametersetName='None')]
    Param(
     [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
     [string]$Identity,
     [Parameter(Mandatory=$true)]
     [ValidateSet("Inbox","Calendar","Contacts","MsgFolderRoot","DeletedItems","Root","Tasks","SentItems")]
     [string]$Folder,
     [Parameter(Mandatory=$false)]
     [string]$SubFolder,
     [Parameter(Mandatory=$false)]
     [string]$Subject,
     [Parameter(Mandatory=$false)]
     [string]$ToAddress,
     [Parameter(Mandatory=$false)]
     [string]$FromAddress,
     [Parameter(Mandatory=$false)]
     [datetime]$StartDate,
     [Parameter(Mandatory=$false)]
     [datetime]$EndDate,
     [Parameter(Mandatory=$false)]
     [string]$Resultsize
    )
    #Internal Function for grabbing mail items
    function getitems([System.Collections.Generic.List[PSObject]]$list, $Resultsize, $folderObj, $searchFilter) {
        $pageSize = 1000
        $offset = 0
        $itemsView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(($pageSize+1),$offset)
        $moreItems = $true
        if($items.Count -eq 0) {
            $list = New-Object System.Collections.Generic.List[PSObject]
            }
        while($moreItems -eq $true -and $list.Count -as [double] -lt $Resultsize) {
            $mailItems = $exService.FindItems($folderObj.Id, $searchFilter, $itemsView)
            foreach($item in $mailItems.Items) {
                $list.Add($item)
                if($list.Count -eq $Resultsize) {
                    break
                    }
                }
            $itemsView.Offset += $pageSize
            $moreItems = $mailItems.MoreAvailable
            }
        return [System.Collections.Generic.List[PSObject]]$list
        }
    #Checking for EWS Service
    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    #Resolving Identity
    $mailboxName = resolveName -Identity $Identity
    if($mailboxName -eq $null) {
        return
    }
    #Setting up search criteria
    $searchFilter = New-Object System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.SearchFilter]
    if($Subject) {
        $subjectItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject
        $subSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($subjectItem, $Subject)
        $searchFilter.Add($subSearch)
        }
    if($ToAddress) {
        $toAddressItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DisplayTo
        $toSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($toAddressItem, $ToAddress)
        $searchFilter.Add($toSearch)
        }
    if($FromAddress) {
        $FromAddressItem = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From
        $fromSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($FromAddressItem, $FromAddress)
        $searchFilter.Add($fromSearch)
        }
    if($StartDate) {
        $startDateItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
        $startDateSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo($startDateItem, $StartDate)
        $searchFilter.Add($startDateSearch)
        }
    if($EndDate) {
        $endDateItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
        $endDateSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo($endDateItem, $EndDate)
        $searchFilter.Add($endDateSearch)
        }
    if($searchFilter.Count -eq 0) {
        $searchFilter = $null
        }
    elseif($searchFilter.Count -eq 1) {
        $searchFilter = $searchFilter[0]
        }
    else {
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection('and', $searchFilter.ToArray())
        }

    #Connecting to mailbox
    $ExService.ImpersonatedUserId = New-Object Microsoft.Exchange.Webservices.Data.ImpersonatedUserID([Microsoft.Exchange.Webservices.Data.ConnectingIDType]::SmtpAddress,$MailboxName)
    $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Folder,$MailboxName)
    try {
        $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderID)
    }
    catch {
        Write-Error $_
        return
    }
    if(!$SubFolder) {
        $subFolders = New-Object System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ServiceObject]
        $pageSize = 1000
        $offset = 0
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(($pageSize+1),$offset)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        if($Folder -like "MsgFolderRoot") {
            $folderClass = [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass
            $folderFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($folderClass,"IPF.Note")
            }
        else {
            $folderFilter = $null
            }
        $moreFolders = $true
        while($moreFolders) {
            $folderItems = $exService.FindFolders($folderObj.Id,$folderFilter,$folderView)
            foreach($fold in $folderItems.Folders) {
                $subfolders.Add($fold)
                }
            $folderView.Offset += $pageSize
            $moreFolders = $folderItems.MoreAvailable
            }
        }
    else {
         $folderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SubFolder)
         $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
         $findFolderResults = $exService.FindFolders($folderObj.Id,$folderSearchFilter,$folderView)
         if($findFolderResults.Folders.Count -lt 1) {
            Write-Error "Subfolder does not exist in current context."
            return 
         }
         else {
             $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService, $findFolderResults.Id)
         }
    }
    
    try { $Resultsize = [int]$Resultsize }
    catch {}

    #Retreving Items
    $items = New-Object System.Collections.Generic.List[PSObject]
    if($Resultsize -eq "Unlimited") {
        $Resultsize = [double]::PositiveInfinity
        }
    elseif($Resultsize -eq 0) {
        $Resultsize = 1000
        }
    $items = getitems -list $items -Resultsize $Resultsize -folderObj $folderObj -searchFilter $searchFilter
    if($subFolders.Count -gt 0 -and $items.Count -as [double] -lt $Resultsize) {
        foreach($sub in $subFolders) {
            $items = getitems -list $items -Resultsize $Resultsize -folderObj $sub -searchFilter $searchFilter
            }
        }

    #Returning mail items found.
    return $items
    }

function Remove-MailItem {
    <#
    .SYNOPSIS
    Removes mail items found by the Get-MailItem. 
    
    .DESCRIPTION
    This cmdlet is used to delete a MailItem from a mailbox. It accepts a MailItem object either through piping or variable.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER MailItem
    This parameter accepts a Mail Item from the Get-MailItem cmdlet. It can be either piped or set up as a variable. 

    .PARAMETER Confirm
    This parameter is used to not require confirmation on each deletion of a message. If it is not flagged, it will show the subject and as for confirmation of delete.

    .PARAMETER DeleteType
    This parameter specifies how you will delete the message, it accepts three parameters, MoveToDeletedItems, SoftDelete, and HardDelete. MoveToDeletedItems moves the message to deleted items, SoftDelete puts the message in the mailbox dumpster, and HardDelete completely removes the message with no way to restore. 

    .EXAMPLE
    Get-MailItem -Identity jsmith -Folder Inbox -Subject "This is totally a virus!" | Remove-MailItem -DeleteType HardDelete -Confirm:$false

    .EXAMPLE
    $mailItem = Get-MailItem -Identity mmuhammad -Folder Calendar -Subject "Meeting that shouldn't exist!" -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date).AddDays(-13)
    Remove-MailItem -MailItem $mailItem -DeleteType SoftDelete

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
    Param(
     [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
     [Microsoft.Exchange.WebServices.Data.Item]$MailItem,
     [Parameter(Mandatory=$false)]
     [bool]$Confirm = $true,
     [Parameter(Mandatory=$true,Position=1)]
     [ValidateSet("HardDelete","SoftDelete","MoveToDeletedItems")]
     [string]$DeleteType
    )

    #Checking for EWS Service
    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteType
    $message = [Microsoft.Exchange.WebServices.Data.Item]::Bind($exService,$MailItem.Id)
    if($Confirm -eq $false) {
        $message.Delete($deleteMode)
    }
    else {
        $answer = Read-Host "Are you sure you want to delete $($message.Subject)? (Y/N)"
        if($answer -like "Y*") {
            $message.Delete($deleteMode)
        }
        else {
            return
        }
    }
}

function Export-MailItem {
    <#
    .SYNOPSIS
    Exports mail items found by the Get-MailItem. 
    
    .DESCRIPTION
    This cmdlet is used to export a mail item, either as a .eml or a .ics. It accepts a mail item either through piping or variable.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER MailItem
    This parameter accepts a Mail Item from the Get-MailItem cmdlet. It can be either piped or set up as a variable. 

    .PARAMETER Path
    This parameter specifies the path where the mail item will be exported, including the name of the mail item. You should specify .eml for an email, .ics for a calendar item, and .vcf for contacts.

    .EXAMPLE
    Get-MailItem -Identity jsmith -Folder Inbox -Subject "This is totally a virus!" | Export-MailItem -Path "c:\users\adminalias\documents\quarantine\virusemail.eml"

    .EXAMPLE
    $mailItem = Get-MailItem -Identity mmuhammad -Folder Calendar -Subject "Meeting that shouldn't exist!" -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date).AddDays(-13)
    Export-MailItem -MailItem $mailItem -Path "C:\scripts\MailItemOutputs\THATMeeting.ics"

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
    Param(
     [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
     [Microsoft.Exchange.WebServices.Data.Item]$MailItem,
     [string]$Path
    )
    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    try {
        $message = [Microsoft.Exchange.WebServices.Data.Item]::Bind($exService,$MailItem.Id)
        }
    catch {
        return $false
        }
    $itemProperties = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
    $message.Load($itemProperties)
    try {
        $fstream = New-Object System.IO.FileStream($Path,[System.IO.FileMode]::Create)
        $fstream.Write($message.MimeContent.Content, 0, $message.MimeContent.Content.Length)
        $fstream.Close()
        }
    catch {
        Write-Host "Unable to write to disk"
        return $false
        }
    return $true
    }

function Import-MailItem {
	<#
    .SYNOPSIS
    Imports mail items exported by Export-MailItem. 
    
    .DESCRIPTION
    This cmdlet is used to import a mail item. They can be either a .eml for an email, a .ics for a calendar event, or .vcf for a contact. 

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER Path
    This parameter specifies the path to the file that is to be imported.
	This is a required parameter.

    .PARAMETER TargetMailbox
    This parameter specifies the target mailbox that will receive the import.
	This is a required parameter.

	.PARAMETER TargetFolder
	This parameter specifies the target folder within the mailbox to import the file to.
    This is a required parameter. 
    
    .PARAMETER SubFolder
    This parameter allows you to specify a subfolder within the well known folder specified in the TargetFolder parameter.

    .EXAMPLE
    Import-Mail -Path ".\goodMeeting.ics" -TargetMailbox "jtaylor" -TargetFolder "Calendar"

    .EXAMPLE
    Import-Mail -Path ".\SuperImportantEmail.eml" -TargetMailbox "rchapman@contoso.com" -TargetFolder "Inbox"

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
	Param(
		[Parameter(Mandatory=$true,Position=0)]
		[String]$Path,
		[Parameter(Mandatory=$true,Position=1)]
		[String]$TargetMailbox,
		[Parameter(Mandatory=$true,Position=2)]
        [ValidateSet("Inbox","Calendar","Contacts","Tasks")]
        [string]$TargetFolder,
        [Parameter(Mandatory=$false)]
        [string]$SubFolder
		)
    if(!$exService) {
            Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
            return
        }
    #Resolving Identity
    $mailboxName = resolveName -Identity $TargetMailbox
    if($mailboxName -eq $null) {
        return
    }
    #Setting up folder access.
    $ExService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxName)
    $folderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$TargetFolder, $mailboxName)
    try {
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderID)
    }
    catch {
        Write-Error $_
        return
    }
    if($SubFolder) {
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $folderSearchFilter = = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SubFolder)
        $findFolderResults = $ExService.FindFolders($folder.ID, $folderSearchFilter, $folderView)
        if($findFolderResults.Folders.Count -eq 1) {
            $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService, $findFolderResults.Id)
        }
    }
    
    $importEmail = Get-Item $Path
    if($TargetFolder -like "Inbox") {
        $uploadEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($exService)
    }
    elseif($TargetFolder -like "Calendar") {
        $uploadEmail = New-Object Microsoft.Exchange.WebServices.Data.Appointment($exService)
    }
    elseif($TargetFolder -like "Contacts") {
        $uploadEmail = New-Object Microsoft.Exchange.WebServices.Data.Contact($exService)
    }
    elseif($TargetFolder -like "Tasks") {
        $uploadEmail = New-Object Microsoft.Exchange.WebServices.Data.Task($exService)
    }
    
    [byte[]]$emailInByte = Get-Content -Encoding Byte $importEmail
    $uploadEmail.MimeContent = New-Object Microsoft.Exchange.Webservices.Data.MimeContent("us-ascii", $emailInByte)
    $PR_Flags = New-Object Microsoft.Exchange.Webservices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $uploadEmail.SetExtendedProperty($PR_Flags,"1")
    if($TargetFolder -like "Calendar") {
        $uploadEmail.Save($folder.ID, [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
    }
    else {
        $uploadEmail.Save($folder.ID)
    }
}
    
function Get-MailFolder {
    <#
    .SYNOPSIS
    Lists subfolders in a given well known folder. 
    
    .DESCRIPTION
    This cmdlet returns a list of folders found within a user's mailbox. It accepts well known folder names such as Inbox, Calendar, Contacts, or Tasks.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER Identity
    This parameter specifies the identity of the mailbox to be searched
	This is a required parameter.

    .PARAMETER Folder
    This parameter specifies the well known folder to be searched. It accepts Inbox, Calendar, Contacts, or Tasks.
	This is a required parameter.

	.PARAMETER SubFolder
	This parameter specifies the name of the sub folder being searched for. If none is specified, it will return all subfolders below the specified well known folder.
    
    .PARAMETER Resultsize
	This parameter specifies the number of folders returned. By default it returns 1000. It also accepts Unlimited.

    .EXAMPLE
    Get-MailFolder -Identity nholmes -Folder Inbox -Resultsize Unlimited

    .EXAMPLE
    Get-MailFolder -Identity cbooth -Folder Calendar

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Identity,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateSet("Inbox","Calendar","Contacts","Tasks")]
        [string]$Folder,
        [Parameter(Mandatory=$false)]
        [string]$SubFolder,
        [Parameter(Mandatory=$false)]
        [string]$Resultsize = "1000"
    )
    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    #Resolving Identity
    $mailboxName = resolveName -Identity $Identity
    if($mailboxName -eq $null) {
        return
    }
    $ExService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxName)
    $folderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Folder, $mailboxName)
    try {
        $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderID)
    }
    catch {
        Write-Error $_
        return
    }
    try {$Resultsize = [int]$Resultsize}
    catch {}
    if($Resultsize -eq "Unlimited") {
        $Resultsize = [double]::PositiveInfinity
    }
    if($SubFolder) {
        $subFolders = New-Object System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ServiceObject]
        $folderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $SubFolder)
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $findFolderResults = $exService.FindFolders($folderObj.Id,$folderSearchFilter,$folderView)
        if($findFolderResults.Folders.Count -lt 1) {
            Write-Error "Subfolder does not exist in current context."
            return
        }
        else {
            $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService, $findFolderResults.Id)
        }
        $subFolders.Add($folderObj)
        $pageSize = 1000
        $offset = 0
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(($pageSize+1),$offset)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $moreFolders = $true
        while($moreFolders -and $subFolders.Count -lt $Resultsize) {
            $folderItems = $exService.FindFolders($folderObj.Id,$folderFilter,$folderView)
            foreach($fold in $folderItems.Folders) {
                if($subFolders.Count -lt $Resultsize) {
                    $subfolders.Add($fold)
                }
                else {
                    break
                }
            }
            $folderView.Offset += $pageSize
            $moreFolders = $folderItems.MoreAvailable
            }
    }
    else {
        $subFolders = New-Object System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ServiceObject]
        $subFolders.Add($folderObj)
        $pageSize = 1000
        $offset = 0
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(($pageSize+1),$offset)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $moreFolders = $true
        while($moreFolders -and $subFolders.Count -lt $Resultsize) {
            $folderItems = $exService.FindFolders($folderObj.Id,$folderFilter,$folderView)
            foreach($fold in $folderItems.Folders) {
                if($subFolders.Count -lt $Resultsize) {
                    $subfolders.Add($fold)
                }
                else {
                    break
                }
            }
            $folderView.Offset += $pageSize
            $moreFolders = $folderItems.MoreAvailable
            }
    }
    #Returning folders
    return $subFolders
}

function New-MailFolder {
    <#
    .SYNOPSIS
    Creates a new subfolder in a given well known folder. 
    
    .DESCRIPTION
    This cmdlet creates a new subfolder in a well known folder's root. It can create folders in Inbox, Calendar, Contacts, or Tasks.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER Identity
    This parameter specifies the identity of the mailbox that will have the folder created.
	This is a required parameter.

    .PARAMETER Folder
    This parameter specifies the well known folder. It accepts Inbox, Calendar, Contacts, or Tasks.
	This is a required parameter.

	.PARAMETER NewSubFolder
    This parameter specifies the name of the folder to be created.
    This is a required parameter
    
    .EXAMPLE
    New-MailFolder -Identity nholmes -Folder Inbox -NewSubFolder "Cool dudes folder"

    .EXAMPLE
    New-MailFolder -Identity cbooth -Folder Calendar -NewSubFolder "A new calendar was made!"

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Identity,
        [Parameter(Mandatory=$true,Position=1)]
        [ValidateSet("Inbox","Calendar","Contacts","Tasks")]
        [string]$Folder,
        [Parameter(Mandatory=$true,Position=2)]
        [string]$NewSubFolder
    )
    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    #Resolving Identity
    $mailboxName = resolveName -Identity $Identity
    if($mailboxName -eq $null) {
        return
    }
    $ExService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $mailboxName)
    $folderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$Folder, $mailboxName)
    try {
        $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderID)
    }
    catch {
        Write-Error $_
        return
    }
    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $folderSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$NewSubFolder)
    $folderResults = $exService.FindFolders($folderObj.Id,$folderSearchFilter,$folderView)
    if($folderResults.Folders.Count -eq 0) {
        $NewFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($exService)
        $NewFolder.DisplayName = $NewSubFolder
        if($Folder -eq "Inbox") {
            $NewFolder.FolderClass = "IPF.Note"
        }
        elseif($Folder -eq "Calendar") {
            $NewFolder.FolderClass = "IPF.Appointment"
        }
        elseif($Folder -eq "Contacts") {
            $NewFolder.FolderClass = "IPF.Contact"
        }
        elseif($Folder -eq "Tasks") {
            $NewFolder.FolderClass = "IPF.Task"
        }
        $NewFolder.Save($folderID)
    }
    else {
        Write-Error "This folder already exists!"
        return
    }
}

function Remove-MailFolder {
    <#
    .SYNOPSIS
    Removes a folder found by the Get-MailFolder cmdlet. 
    
    .DESCRIPTION
    This cmdlet is used to delete a MailFolder from a mailbox. It accepts a MailItem object either through piping or variable. It cannot be used to delete well known folders.

    The Connect-ExWebService must be run before this cmdlet.

    .PARAMETER MailFolder
    This parameter accepts a Mail Folder from the Get-MailFolder cmdlet. It can be either piped or set up as a variable. 

    .PARAMETER Confirm
    This parameter is used to not require confirmation on each deletion of a folder. If it is not flagged, it will show the subject and as for confirmation of delete.

    .PARAMETER DeleteType
    This parameter specifies how you will delete the folder, it accepts three parameters, MoveToDeletedItems, SoftDelete, and HardDelete. MoveToDeletedItems moves the folder to deleted items, SoftDelete puts the folder in the mailbox dumpster, and HardDelete completely removes the folder with no way to restore. 

    .EXAMPLE
    Get-MailItem -Identity jsmith -Folder Inbox -Subject "This is totally a virus!" | Remove-MailItem -DeleteType HardDelete -Confirm:$false

    .EXAMPLE
    $mailItem = Get-MailItem -Identity mmuhammad -Folder Calendar -Subject "Meeting that shouldn't exist!" -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date).AddDays(-13)
    Remove-MailItem -MailItem $mailItem -DeleteType SoftDelete

    .FUNCTIONALITY
    General Cmdlet
    
    #>
    [cmdletbinding()]
    Param(
     [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
     [Microsoft.Exchange.WebServices.Data.Folder]$MailFolder,
     [Parameter(Mandatory=$false)]
     [bool]$Confirm = $true,
     [Parameter(Mandatory=$true,Position=1)]
     [ValidateSet("HardDelete","SoftDelete","MoveToDeletedItems")]
     [string]$DeleteType
    )

    if(!$exService) {
        Write-Error "You are not connected to EWS! Please run the Connect-EWSService cmdlet before running this!"
        return
        }
    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteType
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$MailFolder.Id)
    if($Confirm -eq $false) {
        $folder.Delete($deleteMode)
    }
    else {
        $answer = Read-Host "Are you sure you want to delete $($message.Subject)? (Y/N)"
        if($answer -like "Y*") {
            $folder.Delete($deleteMode)
        }
        else {
            return
        }
    }

}

function resolveName ($Identity) {
    if($Identity -notlike "*@*") {
        $nameResolutionCollection = $ExService.ResolveName($Identity, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly ,$true)
        if($nameResolutionCollection.Count -ne 1) {
            Write-Error "Unable to resolve mailbox by Identity! Please verify that the Identity is correct" -Category InvalidData
            return
        }
        else {
            $mailboxName = $nameResolutionCollection[0].Mailbox.Address
        }
    }
    else {
        $mailboxName = $Identity
    }
    return $mailboxName
}

Export-ModuleMember -Function Get-MailItem, Remove-MailItem, Export-MailItem, Import-MailItem, Get-MailFolder, New-MailFolder, Connect-ExWebService, Remove-MailFolder