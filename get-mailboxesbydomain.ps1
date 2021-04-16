$CSVPath = "C:\temp\mailboxes.csv"
$domain = "@emiindustries.com"

$Mailboxes = (Get-Mailbox).EmailAddresses.SmtpAddress | ? {$_ -match $domain} | Get-Mailbox | select -Unique

$CSV = @()
foreach($Mailbox in $Mailboxes)
    {
        $MailboxStats = (Get-MailboxStatistics $Mailbox)

        if ($mailbox.UseDatabaseQuotaDefaults -eq $true)
            {
                $ProhibitSendReceiveQuota = (Get-MailboxDatabase $mailbox.Database).ProhibitSendReceiveQuota.Value.ToMB()
            }
        if ($mailbox.UseDatabaseQuotaDefaults -eq $false)
            {
                $ProhibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota.Value.ToMB()
            }

        $CSVLine = New-Object System.Object
        $CSVLine | Add-Member -Type NoteProperty -Name DisplayName -Value $Mailbox.DisplayName
        $CSVLine | Add-Member -Type NoteProperty -Name UserName -Value $Mailbox.SamAccountName
        $CSVLine | Add-Member -Type NoteProperty -Name PrimarySMTP -Value $Mailbox.WindowsEmailAddress
        $CSVLine | Add-Member -Type NoteProperty -Name OrganizationalUnit -Value $Mailbox.OrganizationalUnit
        $CSVLine | Add-Member -Type NoteProperty -Name EmailAliases -Value ($Mailbox.EmailAddresses.SmtpAddress -join "; ")
        $CSVLine | Add-Member -Type NoteProperty -Name TotalItemSizeInKB -Value $MailboxStats.TotalItemSize.Value.ToKB()
        $CSVLine | Add-Member -Type NoteProperty -Name ItemCount -Value $MailboxStats.ItemCount
        $CSVLine | Add-Member -Type NoteProperty -Name StorageLimitStatus -Value $Mailbox.StorageLimitStatus
        $CSVLine | Add-Member -Type NoteProperty -Name UseDatabaseQuotaDefaults -Value $Mailbox.UseDatabaseQuotaDefaults
        $CSVLine | Add-Member -Type NoteProperty -Name ProhibitSendReceiveQuotaInMB -Value $ProhibitSendReceiveQuota
        $CSV += $CSVLine
    }

$CSV | Sort TotalItemSize -Descending | Export-Csv -NoTypeInformation $CSVPath