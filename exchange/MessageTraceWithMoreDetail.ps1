function Get-MessageTraceWithMoreDetail
{
<#
.Synopsis
   Trace mail messages with more detail
.DESCRIPTION
   This script traces mail messages and provides more detail exposing each event and outputs a single object for review.
   Start date and end date values are set to 10 days old and current date respectively, but you can overide if required.
.EXAMPLE
   Get-MessageTraceWithMoreDetail -startDate (Get-Date).AddDays(-3) -endDate $endDate = (Get-Date).AddDays(-1)
#>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0)]
        [datetime]$startDate = (Get-Date).AddDays(-10),

        [Parameter(Position=1)]
        [datetime]$endDate = (Get-Date)
    )

try {

        $senderAddress = Read-Host -Prompt "What is the sender address or domain? eg. @domain.com or user@domain.com"

        Clear-Variable -Name recipientCheck,subjectCheck -ErrorAction SilentlyContinue

        while ($recipientCheck -notmatch "Y|N")
        {
            $recipientCheck = Read-Host -Prompt "Do you want to search using a recipient filter? [ Y | N ]"
        }

        while ($subjectCheck -notmatch "Y|N")
        {
            $subjectCheck = Read-Host -Prompt "Do you want to search using a subject filter? [ Y | N ]"
        }

        if ($subjectCheck -eq "Y" -and $recipientCheck -eq "Y") {
        
                $recipientFilter = Read-Host -Prompt "What is the recipients email address? eg. user@domain.com"
                $subjectFilter = Read-Host -Prompt "What words does the subject contain?"

                $messagesToReview = Get-MessageTrace -StartDate $startDate -EndDate $endDate -SenderAddress $senderAddress -RecipientAddress $recipientFilter | Where-Object -FilterScript {$_.Subject -like "*$subjectFilter*"}

        }
        elseif ($subjectCheck -eq "Y" -and $recipientCheck -eq "N") {

                $subjectFilter = Read-Host -Prompt "What words does the subject contain?"

                $messagesToReview = Get-MessageTrace -StartDate $startDate -EndDate $endDate -SenderAddress $senderAddress | Where-Object -FilterScript {$_.Subject -like "*$subjectFilter*"}

        }
        elseif ($subjectCheck -eq "N" -and $recipientCheck -eq "Y") {

                $recipientFilter = Read-Host -Prompt "What is the recipients email address? eg. user@domain.com"

                $messagesToReview = Get-MessageTrace -StartDate $startDate -EndDate $endDate -SenderAddress $senderAddress -RecipientAddress $recipientFilter

        }
        else {

               $messagesToReview = Get-MessageTrace -StartDate $startDate -EndDate $endDate -SenderAddress $senderAddress

        }

    Write-Host ""
    Write-Host ""
    Write-Host "============ Message Report $(get-date) ============" -ForegroundColor Cyan

    foreach ($message in $messagesToReview) {

                # Convert to local time
                $messageReceivedDate = Get-LocalTime -UTCTime $message.Received


                $customMessageObjectProps = [ordered]@{
                    'Step 0 : Sender IP' = $message.FromIP ;
                    'Step 0 : From address' = $message.SenderAddress ;
                    'Step 0 : Date received' = $messageReceivedDate ;
                    'Step 0 : To address' = $message.RecipientAddress ;
                    'Step 0 : Message subject' = $message.Subject ;
                    'Step 0 : Message status' = $message.Status ;
                    'Step 0 : Message size (KB)' = $([math]::Round(($message.Size / 1KB),2)) ;
                    'Step 0 : Message ID' = $message.MessageId}

                $customMessageObject = ""

                $customMessageObject = New-Object -TypeName PSObject -Property $customMessageObjectProps

                $messageDetail = $message | Get-MessageTraceDetail

                [int]$c = ""

                foreach ($event in $messageDetail) {

                    $xml = [xml]$event.Data
                    $eventReport = $xml.root.MEP

                    $c++
                    
                    # Convert to local time
                    $adjustedDate = Get-LocalTime -UTCTime $event.Date
             
                    $customMessageObject | Add-Member -NotePropertyName "Step $($c) : Action taken:" -NotePropertyValue $($event.Detail)
                    $customMessageObject | Add-Member -NotePropertyName "Step $($c) : Action time:" -NotePropertyValue $($adjustedDate.ToString())
            

                        foreach ($xmlProp in $eventReport) {

                            if ( 'string' -in (($eventReport | Get-Member).Name)) {
                
                                   if ($xmlProp.String -ne $null) {

                                        if ($xmlProp.Name -ne "RecipientReference") {
                        
                                        $customMessageObject | Add-Member -NotePropertyName "Step $($c) : $($xmlProp.Name)" -NotePropertyValue $xmlProp.string -Force

                                        }
                       
                                   }           
                
                            }
                            elseif ( 'integer' -in (($eventReport | Get-Member).Name)) {
                
                                   if ($xmlProp.integer -ne $null) {
                       
                                        $customMessageObject | Add-Member -NotePropertyName "Step $($c) : $($xmlProp.Name)" -NotePropertyValue $xmlProp.integer -Force
                       
                                   }           
                
                            }
                        }
                
                   }

        Write-Output $customMessageObject

        }

    }
    catch {

    Write-Host "Failed to perform message trace with more detail"
    Write-Host "$($_)"
    Write-Host "Line Number: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Host "Offset: $($_.InvocationInfo.OffsetInLine)"
    Write-Host "Line: $($_.InvocationInfo.Line)"

    }

}


function Get-LocalTime($UTCTime)
{
$strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName
$TZ = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone)
$LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
Return $LocalTime
}