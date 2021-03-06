$Mailbox1 = read-host -Prompt 'Input email address you need to remove access from'
$Mailbox2 = read-host -Prompt 'Input email address that no longer needs access'

Remove-MailboxPermission -Identity $Mailbox1 -User $Mailbox2 -AccessRights 'FullAccess' -InheritanceType All
write-host "Access to '$mailbox1' by '$mailbox2' has been removed"