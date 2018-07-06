#
# Create-Folders.ps1
#

param (
	[Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the mailbox to be accessed")]
	[ValidateNotNullOrEmpty()]
	[string]$Mailbox,
	
	[Parameter(Position=1,Mandatory=$True,HelpMessage="Specifies the folder(s) that should be checked/created.  For multiple folders, separate using semicolon")]
	[ValidateNotNullOrEmpty()]
	[string]$RequiredFolders,
	
	[Parameter(Mandatory=$False,HelpMessage="The folder that should contain the subfolders (default is Inbox)")]
	[string]$ParentFolder,
	
	[Parameter(Mandatory=$False,HelpMessage="Username used to authenticate with EWS")]
	[string]$AuthUsername,
	
	[Parameter(Mandatory=$False,HelpMessage="Password used to authenticate with EWS")]
	[string]$AuthPassword,
	
	[Parameter(Mandatory=$False,HelpMessage="Domain used to authenticate with EWS")]
	[string]$AuthDomain,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether we are using impersonation to access the mailbox")]
	[switch]$Impersonate,
	
	[Parameter(Mandatory=$False,HelpMessage="EWS Url (if omitted, then autodiscover is used)")]	
	[string]$EwsUrl,
	
	[Parameter(Mandatory=$False,HelpMessage="Path to managed API (if omitted, a search of standard paths is performed)")]	
	[string]$EWSManagedApiPath = $Env:ProgramFiles + "\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll",
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to ignore any SSL errors (e.g. invalid certificate)")]	
	[switch]$IgnoreSSLCertificate,
	
	[Parameter(Mandatory=$False,HelpMessage="Whether to allow insecure redirects when performing autodiscover")]	
	[switch]$AllowInsecureRedirection,

	[Parameter(Mandatory=$False,HelpMessage="If specified, no changes will be applied")]	
	[switch]$WhatIf

)


Function SearchDll()
{
	# Search for a program/library within Program Files (x64 and x86)
	$path = $args[0]
	$programDir = $Env:ProgramFiles
	if (Get-Item -Path ($programDir + $path) -ErrorAction SilentlyContinue)
	{
		return $programDir + $path
	}
	
	$programDir = [environment]::GetEnvironmentVariable("ProgramFiles(x86)")
	if ( [string]::IsNullOrEmpty($programDir) ) { return "" }
	
	if (Get-Item -Path ($programDir + $path) -ErrorAction SilentlyContinue)
	{
		return $programDir + $path
	}
}

Function LoadEWSManagedAPI()
{
	# Check EWS Managed API available
	
	if ( !(Get-Item -Path $EWSManagedApiPath -ErrorAction SilentlyContinue) )
	{
		$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll")
		if ( [string]::IsNullOrEmpty($EWSManagedApiPath) )
		{
			$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll")
			if ( [string]::IsNullOrEmpty($EWSManagedApiPath) )
			{
				$EWSManagedApiPath = SearchDll("\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll")
			}
		}
	}
	
	If ( ![string]::IsNullOrEmpty($EWSManagedApiPath) )
	{
		# Load EWS Managed API
		Write-Verbose ([string]::Format("Using managed API found at: {0}", $EWSManagedApiPath))
		Add-Type -Path $EWSManagedApiPath
		return $true
	}
	return $false
}

Function GetFolder()
{
	# Return a reference to a folder specified by path
	
	$RootFolder, $FolderPath = $args[0]
	
	$Folder = $RootFolder
	if ($FolderPath -ne '\')
	{
		$PathElements = $FolderPath -split '\\'
		For ($i=0; $i -lt $PathElements.Count; $i++)
		{
			if ($PathElements[$i])
			{
				$View = New-Object  Microsoft.Exchange.WebServices.Data.FolderView(2,0)
				$View.PropertySet = [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
						
				$SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $PathElements[$i])
				
				$FolderResults = $Folder.FindFolders($SearchFilter, $View)
				if ($FolderResults.TotalCount -ne 1)
				{
					# We have either none or more than one folder returned... Either way, we can't continue
					$Folder = $null
					Write-Verbose ([string]::Format("Failed to find {0}", $PathElements[$i]))
					Write-Verbose ([string]::Format("Requested folder path: {0}", $FolderPath))
					break
				}
				
				$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderResults.Folders[0].Id)
			}
		}
	}
	
	return $Folder
}

Function CreateFolders()
{
	$FolderId = $args[0]
	Write-Verbose "Binding to folder with id $FolderId"
	$folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderId)
	if (!$folder) { return }
	
	foreach ($requiredFolder in $FolderCheckList)
	{
		Write-Verbose "Checking for existence of $requiredFolder"
		$rf = GetFolder( $folder, $requiredFolder )
		if ( $rf )
		{
			Write-Host "$requiredFolder already exists" -ForegroundColor Green
		}
		Else
		{
			# Create the folder
			if (!$WhatIf)
			{
				$rf = New-Object Microsoft.Exchange.WebServices.Data.Folder($service)
				$rf.DisplayName = $requiredFolder
				$rf.Save($FolderId)
				if ($rf.Id.UniqueId)
				{
					Write-Host "$requiredFolder created successfully" -ForegroundColor Green
				}
			}
			Else
			{
				Write-Host "$requiredFolder would be created" -ForegroundColor Yellow
			}
		}
	}
}

Function ProcessMailbox()
{
	# Process mailbox
	
	Write-Host "Processing mailbox $Mailbox" -ForegroundColor White
	if ( $WhatIf )
	{
		Write-Host "NO CHANGES WILL BE APPLIED" -ForegroundColor Red
	}

	# Set EWS URL if specified, or use autodiscover if no URL specified.
	if ($EwsUrl)
	{
		$service.URL = New-Object Uri($EwsUrl)
	}
	else
	{
		Write-Verbose "Performing autodiscover for $Mailbox"
		if ( $AllowInsecureRedirection )
		{
			$service.AutodiscoverUrl($Mailbox, {$True})
		}
		else
		{
			$service.AutodiscoverUrl($Mailbox)
		}
		if ([string]::IsNullOrEmpty($service.Url))
		{
			Write-Host "Autodiscover failed, cannot process mailbox" -ForegroundColor Red
			return
		}
		Write-Verbose ([string]::Format("EWS Url found: {0}", $service.Url))
	}
	 
	# Set impersonation if specified
	if ($Impersonate)
	{
		Write-Verbose "Impersonating $Mailbox"
		$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox)
		$FolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
	}
	else
	{
		# If we're not impersonating, we will specify the mailbox in case we are accessing a mailbox that is not the authenticating account's
		$mbx = New-Object Microsoft.Exchange.WebServices.Data.Mailbox( $Mailbox )
		$FolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $mbx )
	}
	
	if ($ParentFolder)
	{
		$Folder = GetFolder([Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$FolderId), $ParentFolder)
		if (!$FolderId)
		{
			Write-Host "Failed to find folder $ParentFolder" -ForegroundColor Red
			return
		}
		$FolderId = $Folder.Id
	}

	CreateFolders $FolderId
}


# The following is the main script


if (!(LoadEWSManagedAPI))
{
	Write-Host "Failed to locate EWS Managed API, cannot continue" -ForegroundColor Red
	Exit
}


$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

# If we are ignoring any SSL errors, set up a callback
if ($IgnoreSSLCertificate)
{
	Write-Host "WARNING: Ignoring any SSL certificate errors" -ForegroundColor Yellow
	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
}

# Set credentials if specified, or use logged on user.
 if ($AuthUsername -and $AuthPassword)
 {
	Write-Verbose "Applying given credentials for", $AuthUsername
	if ($AuthDomain)
	{
		$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername,$AuthPassword,$AuthDomain)
	} else {
		$service.Credentials = New-Object  Microsoft.Exchange.WebServices.Data.WebCredentials($AuthUsername,$AuthPassword)
	}

} else {
	Write-Verbose "Using default credentials"
    $service.UseDefaultCredentials = $true
}

if ($RequiredFolders.Contains(";"))
{
	# Have more than one folder to check, so convert to array
	$FolderCheckList = $RequiredFolders -split ';'
}
else
{
	$FolderCheckList = $RequiredFolders
}


# Check whether we have a CSV file as input...
$FileExists = Test-Path $Mailbox
If ( $FileExists )
{
	# We have a CSV to process
	$csv = Import-CSV $Mailbox
	foreach ($entry in $csv)
	{
		$Mailbox = $entry.PrimarySmtpAddress
		if ( [string]::IsNullOrEmpty($Mailbox) -eq $False )
		{
			ProcessMailbox
		}
	}
}
Else
{
	# Process as single mailbox
	ProcessMailbox
}