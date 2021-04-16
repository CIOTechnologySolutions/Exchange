Import-Csv C:\Users\kelliott\Documents\customers\EMS\distributiongroups.csv | ForEach-Object{
    $RecipientTypeDetails=$_.RecipientTypeDetails
    $Name = $($_.Name -replace '\s','')[0..63] -join "" # remove spaces first, then truncate to first 64 characters
    $Alias = $($_.Alias -replace '\s','')[0..63] -join "" # remove spaces first, then truncate to first 64 characters
    $DisplayName=$_.DisplayName
    $smtp=$_.PrimarySmtpAddress
    $RequireSenderAuthenticationEnabled=[System.Convert]::ToBoolean($_.RequireSenderAuthenticationEnabled)
    $join=$_.MemberJoinRestriction
    $depart=$_.MemberDepartRestriction
    $ManagedBy=$_.ManagedBy -split ';'
    $AcceptMessagesOnlyFrom=$_.AcceptMessagesOnlyFrom -split ';'
    $AcceptMessagesOnlyFromDLMembers=$_.AcceptMessagesOnlyFromDLMembers -split ';'
    $AcceptMessagesOnlyFromSendersOrMembers=$_.AcceptMessagesOnlyFromSendersOrMembers -split ';'
    
    Write-Output ""
    Write-Output "working on Group: $Name"
    Write-Output ""

    if ($RecipientTypeDetails -eq "MailUniversalSecurityGroup")
        {
        if ($ManagedBy)
            {
            New-DistributionGroup -Type security -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart -ManagedBy $ManagedBy
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
            Else
            {
            New-DistributionGroup -Type security -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart 
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
        }

    if ($RecipientTypeDetails -eq "MailUniversalDistributionGroup")
        {
        if ($ManagedBy)
            {
            New-DistributionGroup -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart -ManagedBy $ManagedBy
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
            Else
            {
            New-DistributionGroup -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart 
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
        }

    if ($RecipientTypeDetails -eq "RoomList")
        {
        if ($ManagedBy)
            {
            New-DistributionGroup -RoomList -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart -ManagedBy $ManagedBy
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
            Else
            {
            New-DistributionGroup -RoomList -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled -MemberJoinRestriction $join -MemberDepartRestriction $depart 
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $Name -HiddenFromAddressListsEnabled $false
            }
        }


    if ($AcceptMessagesOnlyFrom) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFrom $AcceptMessagesOnlyFrom}
    if ($AcceptMessagesOnlyFromDLMembers) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFromDLMembers $AcceptMessagesOnlyFromDLMembers}
    if ($AcceptMessagesOnlyFromSendersOrMembers) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFromSendersOrMembers $AcceptMessagesOnlyFromSendersOrMembers}
  }