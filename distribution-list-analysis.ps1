<#
Uses Get-DistributionGroup cmdlet to get all distribution groups of the Exchange Online tenant

Loops through each DL, finds all members and extracts the data into a table for analysis, exporting to CSV

Attributes collected per record:
- Display name of DL
- Primary SMTP address of DL
- Display name of member
- Primary SMTP address of member

#>

Import-Module ExchangeOnlineManagement

Connect-ExchangeOnline

$ExportPath = "."

$results = @()

$distribution_groups = Get-DistributionGroup -ResultSize Unlimited

foreach ($group in $distribution_groups) {  
    $members = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited

    foreach ($member in $members) {
        $results += [PSCustomObject]@{
            GroupName = $group.DisplayName
            GroupEmail = $group.PrimarySmtpAddress
            MemberName = $member.DisplayName
            MemberEmail = $member.PrimarySmtpAddress
        }
    }
}

$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Disconnect-ExchangeOnline