$Mailbox1 = read-host -Prompt 'Input email address you need access to'
$Mailbox2 = read-host -Prompt 'Input email address to grant rights to'

Add-MailboxPermission -Identity $Mailbox1 -User $Mailbox2 -AccessRights 'FullAccess' 
write-host "Access to '$mailbox1' has been granted to '$mailbox2'"