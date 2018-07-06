<#
Create-Outlook-Signature.ps1
Description: This script auto generates outlook signatures by modifying a prepared template
By Kyle Elliott
kelliott@ciotech.us

Use: 
1. Create a *.docx template file
2. Edit the file within word so it looks the way you want the signature to look and edit text to correspond with the AD properties
further down in the script. For example the test "Title" corresponds to the AD property "title".
3. Edit the section #Custom variables as required (intial setup per company)
4. Run the script as the user who needs the signature made

#>

#Custom variables
$SignatureName = '' #Insert desired name of signature. This will name will appear in Outlook.
$SigSource = "" #Provide full path to the *.docx signature file
$SignatureVersion = "1.0" #Change this if you have updated the signature. If you do not change it, the script will quit after checking if signature is up to date.
$ForceSignature = '0' #If set to '0', the signature will be editable in Outlook and if set to '1' will be non-editable and forced.
$DefaultAddress = ''
$DefaultCity = ''
$DefaultTelephone = '727-232-8356'
 
 
#Environment variables
$AppData=(Get-Item env:appdata).value
$SigPath = '\Microsoft\Signatures'
$LocalSignaturePath = $AppData+$SigPath
$RemoteSignaturePathFull = $SigSource
 
#Copy version file
If (!(Test-Path -Path $LocalSignaturePath\$SignatureVersion))
{
New-Item -Path $LocalSignaturePath\$SignatureVersion -Type Directory
}
Elseif (Test-Path -Path $LocalSignaturePath\$SignatureVersion)
{
Write-Output "Latest signature already exists"
break
}
 
#Check signature path (needs to be created if a signature has never been created for the profile
if (!(Test-Path -path $LocalSignaturePath)) {
	New-Item $LocalSignaturePath -Type Directory
}
 
#Get Active Directory information for current user
$UserName = $env:username
$Filter = "(&amp;(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$ADDisplayName = $ADUser.DisplayName
$ADEmailAddress = $ADUser.mail
$ADTitle = $ADUser.title
$ADDescription = $ADUser.description
$ADTelePhoneNumber = $ADUser.TelephoneNumber
$ADFax = $ADUser.facsimileTelephoneNumber
$ADMobile = $ADUser.mobile
$ADStreetAddress = $ADUser.streetaddress
$ADCity = $ADUser.l
$ADPOBox = $ADUser.postofficebox
$ADCustomAttribute1 = $ADUser.extensionAttribute1
$ADModify = $ADUser.whenChanged
 
#Copy signature templates from domain to local Signature-folder
Write-Output "Copying Signatures"
Copy-Item "$Sigsource" $LocalSignaturePath -Recurse -Force
$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $False
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False
 
#Insert variables from Active Directory to rtf signature-file
$MSWord = New-Object -ComObject word.application
$fullPath = $LocalSignaturePath+'\'+$SignatureName+'.docx'
$MSWord.Documents.Open($fullPath)
 
#User Name $ Designation 
$FindText = "DisplayName" 
$Designation = $ADCustomAttribute1.ToString() #designations in Exchange custom attribute 1
If ($Designation -ne '') { 
	$Name = $ADDisplayName.ToString()
	$ReplaceText = $Name+', '+$Designation
	}
Else {
	$ReplaceText = $ADDisplayName.ToString() 
}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
$LogInfo += $NL+'Username: '+$ReplaceText	
 
#Title		
$FindText = "Title"
$ReplaceText = $ADTitle.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
$LogInfo += $NL+'Title: '+$ReplaceText
	
#Description
 
If ($ADDescription -ne '') { 
   	$FindText = "Description"
   	$ReplaceText = $ADDescription.ToString()
                           }
Else {
	$FindText = " | Description "
   	$ReplaceText = "".ToString()
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Description: '+$ReplaceText
   	
#Street Address
 
If ($ADStreetAddress -ne '') { 
    $FindText = "StreetAddress"
    $ReplaceText = $ADStreetAddress.ToString()
}
Else {
    $FindText = "StreetAddress"
    $ReplaceText = $DefaultAddress
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Street Address: '+$ReplaceText
 
#City
 
If ($ADCity -ne '') { 
    $FindText = "City"
    $ReplaceText = $ADCity.ToString()
}
Else {
    $FindText = "City"
    $ReplaceText = $DefaultCity 
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'City: '+$ReplaceText
	
#Telephone
If ($ADTelephoneNumber -ne "") { 
	$FindText = "TelephoneNumber"
	$ReplaceText = $ADTelephoneNumber.ToString()
}
Else {
    $FindText = "TelephoneNumber"
	$ReplaceText = $DefaultTelephone
}
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Telephone: '+$ReplaceText
	
#Mobile
If ($ADMobile -ne "") { 
	$FindText = "MobileNumber"
	$ReplaceText = $ADMobile.ToString()
   }
Else {
	$FindText = "| Mob MobileNumber "
    $ReplaceText = "".ToString()
	}
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
    $LogInfo += $NL+'MobileNumber: '+$ReplaceText
    
#Email Address
If ($ADEmailAddress -ne "") {
    $FindText = "EmailAddress"
    $ReplaceText = $ADEmailAddress.ToString()
}
 
#Save new message signature 
 
Write-Output "Saving Signatures"
 
#Save HTML
 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
$path = $LocalSignaturePath+'\'+$SignatureName+".htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)
    
#Save RTF 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
$path = $LocalSignaturePath+'\'+$SignatureName+".rtf"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)
	
#Save TXT    
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");
$path = $LocalSignaturePath+'\'+$SignatureName+".txt"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
$MSWord.ActiveDocument.Close()
$MSWord.Quit()
	
 
#Office 2010
If (Test-Path HKCU:'\Software\Microsoft\Office\14.0')
{
Write-Output "Setting signature for Office 2010"
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -ErrorAction SilentlyContinue) 
    {
    Write-Output "Signature already exists"
    } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force
    }
}
If ((Test-Path HKCU:'\Software\Microsoft\Office\14.0') -eq $False)
{
Write-Output "Office 2010 is not installed"
}
#Office 2013 
 
If (Test-Path HKCU:'\Software\Microsoft\Office\15.0')
 
{
Write-Output "Setting signature for Office 2013"
 
If ($ForceSignature -eq '0')
 
{
Write-Output "Setting signature for Office 2013 as available"
 
 
$Outlook = "Outlook"
if ($Outlook -ne $null)
{
Stop-Process -Name $Outlook -Force
}
 
$MSWord = New-Object -comobject word.application
$EmailOptions = $MSWord.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature="$SignatureName"
$EmailSignature.ReplyMessageSignature="$SignatureName"
Stop-Process -Name $Outlook
 
}
 
If ($ForceSignature -eq '1')
{
Write-Output "Setting signature for Office 2013 as forced"
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { 
    New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
}
 
}
 
If ((Test-Path HKCU:'\Software\Microsoft\Office\15.0') -eq $False)
{
Write-Output "Office 2013 is not installed"
}