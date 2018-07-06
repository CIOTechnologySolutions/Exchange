#AddMailboxPermissions.ps1
<#
.Synopsis
    Set full permissions on a mailbox
.parameter target
      Enter the mailbox to be modified
.Parameter user
      Enter the user that will need access to the target mailbox
.Example
      AddMailboxPermissions -target Jim -user frank
.NOTES
      Created to help guide a lower level tech through how to add full access
      Provided without any warranty and entirely as is.
      By Kyle Elliott
      kelliott(at)ciotech(dot)us
#>
{
param(
  [Parameter(Mandatory=$true)]
  [string]$target,
  [Parameter(Mandatory=$True)]
  [string]$user
)

Add-MailboxPermission -Identity $target -User $user -AccessRights 'FullAccess'
write-host "Access to '$target' has been granted to '$user'"
}
