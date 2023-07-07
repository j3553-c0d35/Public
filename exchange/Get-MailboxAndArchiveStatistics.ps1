# Establish output path
$OutputPath = "$env:TEMP\$(Get-date -Format 'yyyyMMddhhmmss')_mailbox_report.csv"

# Establish result array variable
$Result=@()

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Get total mailboxes and establish counter variable
$totalmbx = $mailboxes.Count
$i = 0 

# Loop through each mailbox and perform actions
$mailboxes | ForEach-Object {
    # Increment counter
    $i++
    # Add current mailbox to $mbx variable
    $mbx = $_
    # Reset variables for next loop
    $mba = $null
    $mbs = $null
    $mbasize = $null
    $mbssize = $null
    $MailboxAllocationInGB = $null
    $ArchiveAllocationInGB = $null
    
    # Write progress to host
    Write-Host "Processing $mbx" "$i out of $totalmbx completed"
    
    # Check if archive enabled, if so, get archive stats
    if ($mbx.ArchiveName){
        $mba = Get-MailboxStatistics -Archive $mbx.UserPrincipalName
        
        # Format archive size to GB with 2 decimal places
        if ($mba.TotalItemSize -ne $null){
            $mbasize = [math]::Round(($mba.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
            }
            else{
            $mbasize = 0 
        } 
    }

    # Get mailbox stats
    $mbs = Get-MailboxStatistics $mbx.UserPrincipalName
        
        # Format mailbox size to GB with 2 decimal places
        if ($mbs.TotalItemSize -ne $null){
            $mbssize = [math]::Round(($mbs.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
            }
            else{
            $mbssize = 0 
        } 

    # Get archive allocation (quota) and trim everything but the size in GB
    if ($mbx.ArchiveName){
        $ArchiveAllocationInGB = $mbx.ArchiveQuota.Split('G')
        $ArchiveAllocationInGB = $ArchiveAllocationInGB[0]
    }
    # Get mailbox allocation (quota) and trim everything but the size in GB
    $MailboxAllocationInGB = $mbx.ProhibitSendReceiveQuota.Split('G')
    $MailboxAllocationInGB = $MailboxAllocationInGB[0]

    # Create PSObject and store all relevant information for export
    $Result += New-Object -TypeName PSObject -Property $([ordered]@{ 
        UserName = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        MailboxType = $mbx.RecipientTypeDetails
        MailboxAllocationInGB = $MailboxAllocationInGB
        MailboxSizeInGB = $mbssize
        MailboxItemCount = if ($mbs.ItemCount) {$mbs.ItemCount} Else { $null}
        ArchiveEnabled = if ($mbx.ArchiveName) {"Enabled"} Else { "Disabled"}
        ArchiveName = $mbx.ArchiveName
        ArchiveAllocationInGB = if ($mbx.ArchiveName) {$ArchiveAllocationInGB} Else { $null} 
        ArchiveSizeInGB = $mbasize
        ArchiveItemCount = if ($mba.ItemCount) {$mba.ItemCount} Else { $null}
        AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
    })
}
# Export results to CSV
$Result | Export-CSV $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Output csv file is located here: `n `n $OutputPath `n" -ForegroundColor Yellow