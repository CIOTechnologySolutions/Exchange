{
Param(
    [Parameter(Mandatory=$true)]
    [string]$csv
) #end Param

$users = Import-Csv $csv
foreach ($user in $users) {
    Set-ADUser -Identity $user.SamAccountName -EmailAddress $user.EmailAddress
}
}