<#Windows PowerShell Code
.description
Office365LicenseChange.ps1 - Mass changes office 365 licensing on a customer
.Synopsis
To be used to change licenses in bulk on 365. There are required items that you must install
.Example
.\Office365LicenseChange.ps1
.NOTES
First - Install Microsoft Online Services Sign-In Assistant for IT Professionals (Found here: https://www.microsoft.com/en-us/download/details.aspx?id=41950)
Second - Run from an admin powershell the following command and say yes to any prompt that comes from it: Install-Module MSOnline (ref: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell)
Then, get the right license level by running MsolAccountSku
Enter the proper license info as retrieved from MOSLAccountSku in $oldLicense and $newLicense
Once you do that, run the script and log in to the prompt as an admin on that 365 tenant (Note: You cannot use a partner account for this function)
Known license levels: 
E1: STANDARDPACK
E3: ENTERPRISEPACK
E4: ENTERPRISEWITHSCAL
E5: ENTERPRISEPREMIUM

  By Kyle Elliott
  kelliott(at)ciotech(dot)us
  Provided as is, without warranty
  #>

Import-Module MSOnline
$cred = Get-Credential
Connect-MsolService -Credential $cred

$oldLicense = "365DOMAIN:STANDARDPACK"
$newLicense = "365DOMAIN:ENTERPRISEWITHSCAL"

$users = Get-MsolUser -MaxResults 5000 | Where-Object { $_.isLicensed -eq "TRUE" }

foreach ($user in $users){
    $upn = $user.UserPrincipalName
    foreach ($license in $user.Licenses) {
        if ($license.AccountSkuId -eq $oldLicense) {
            $disabledPlans = @()
            Write-Host("User $upn will go from $oldLicense to $newLicense and will have no options disabled.")
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $newLicense -RemoveLicenses $oldLicense
        }
    }
}
