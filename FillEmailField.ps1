<#Windows PowerShell Code
.description
Scipt to fill out the email field in AD from CSV - FillEmailField.ps1
.Synopsis
This script runs through your CSV, and fills out the email address field on any user listed
CSV must contain the following headers:
sAMAccountName,EmailAddress
.Parameter CSV
Provide path to CSV File that contains a header of sAMAccountName
.Example
FillEmailField.ps1 -csv c:\temp\emailaddress.csv
.NOTES
  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
  #>
﻿
Param(
    [Parameter(Mandatory=$true)]
    [string]$csv
) #end Param

$users = Import-Csv $csv
foreach ($user in $users) {
    Set-ADUser -Identity $user.SamAccountName -EmailAddress $user.EmailAddress
}
