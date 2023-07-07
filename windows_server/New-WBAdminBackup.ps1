function New-WBAdminBackup{
    <#
    .Synopsis
       Gathers information from admin about backup being performed
    .DESCRIPTION
       Obtains input for backup source, destination and exlusions.
    .PARAMETER BackupSource
        Source volumes to backup, only fixed volumes will be valid choices. Baremetal, SystemState and boot volumes will be auto-selected.
    .PARAMETER BackupDestination
        Destination volumes to backups to, multiple volumes can be chosen as targets.
    .PARAMETER Exclude
        Folders and or files to exclude from backup.
    .EXAMPLE
       Get-BackupInfo -BackupSource C:,E: -BackupDestination F: -Exclude C:\Windows\Temp
    #>
    [CmdletBinding()]
    
        param (
            [Parameter(Mandatory=$True,
                       HelpMessage="Enter a drive letter, eg. 'C:\','E:\'")]
                       [alias("BS")]
                       [string[]]$BackupSource,
    
            [Parameter(Mandatory=$True,
                       HelpMessage="Enter a drive letter, eg. 'F','G'")]
                       [alias("BD")]
                       [ValidateScript({if ((Get-ChildItem -Path Env:\SystemDrive).value -like ($_)) {
                            Write-Error -Message "Can not backup to system drive" -ErrorAction Stop}})]
                       [ValidateLength(1)]
                       [ValidatePattern("[A-Z]")]
                       [string[]]$BackupDestination,
    
            [Parameter(HelpMessage="Enter a file or path to exclude, eg. 'C:\Windows\Temp','*.bak'")]
                       [string[]]$Exclude,
    
            [Parameter(HelpMessage="Enter a log file location, eg. 'C:\Backups\Logs\Logfile.log'")]
                       [string]$LogFileLocation,
    
            [Parameter(HelpMessage="Enter an SMTP server, eg. 'mail.company.com.au' or '10.10.10.40'")]
                       [string]$SMTPServer,
    
            [Parameter(HelpMessage="Enter a mail sender, eg. 'backup@company.com.au'")]
                       [string]$MailSender,
    
            [Parameter(HelpMessage="Enter a mail recipient, eg. 'support@myitcompany.com.au'")]
                       [string]$MailRecipient,
    
            [Parameter(HelpMessage="Enter a time to schedule daily backups, eg. '21:00'")]
                       [datetime]$BackupTime
    
            )
    
            
        $DestinationGUID = foreach ($BackupDriveLetter in $BackupDestination)
                                                { (Get-Volume $BackupDriveLetter).ObjectId }
    
        foreach ($GUID in $DestinationGUID) 
                    { if (Get-Volume -ObjectId $GUID)
                        { WBADMIN START BACKUP -backupTarget:$GUID -include:$BackupSource -allCritical -systemState -exclude:$Exclude -vssFull -quiet }}
                            else { Write-Error -Message "Backup target not found!" -ErrorAction Stop }
    
    
        }
    