{
Param(
    [Parameter(Mandatory=$true)]
    [string]$csv
) #end Param

Import-Csv $csv | foreach {
$UPN=$_.UserPrincipalName
$Proxy=$_.ProxyAddresses
Get-ADuser -filter {UserPrincipalName -Eq $UPN} |Set-ADUser -Add @{Proxyaddresses=$proxy}
}
}