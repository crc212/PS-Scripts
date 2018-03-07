#Specify the sending email address and subject. Both are required
$sender = "spammer@nowhere.com"
$subject = "View Your Invoice"

#How many hours back to check
$hours = "2"

#########################################################################
$title = "Message Search"
$message = "Do you want to search for replies and forwards as well? (This will extend the time of the search)"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
switch ($result)
    {
        0 {$rfcheck = "Y"}
        1 {$rfcheck = "N"}
    }
$targetMailbox = "virus@wrberkley.com"
[string]$date = Get-Date -Format yyyy.MM.dd
$targetFolderSearch = $date + " Message Search"
$targetFolderDelete = $date + " Message Delete"
$recipientlist = @()
$table = @()
$rftable = @()
$resultsTable = @()
[int]$resultsCount = "0"
[int]$RFResultsCount = "0"
[int]$resultsCountMB = "0"
[string]$searchQuery = "(From:$sender) AND (Subject:`"$subject`")"
[string]$rfSearchQuery = "(Subject:`"RE: $subject`") OR (Subject:`"FW: $subject`")"

#Search through the tracking logs for the attack
$track = Get-TransportService | Get-MessageTrackingLog -Start (get-date).AddHours(-$hours) -End (Get-Date) -MessageSubject $subject -Sender $sender -EventId DELIVER -ResultSize unlimited
if ($rfcheck -eq "Y"){
    $replytrack = Get-TransportService | Get-MessageTrackingLog -Start (get-date).AddHours(-$hours) -End (Get-Date) -MessageSubject "RE: $subject" -EventId DELIVER -ResultSize unlimited
    $forwardtrack = Get-TransportService | Get-MessageTrackingLog -Start (get-date).AddHours(-$hours) -End (Get-Date) -MessageSubject "FW: $subject" -EventId DELIVER -ResultSize unlimited
}
#Build a table of everyone who received the attack message
foreach ($result in $track){
[string]$recipient = $result.Recipients
$obj = New-Object -TypeName PSObject
$obj | Add-Member -MemberType NoteProperty -Name Recipient -Value $recipient
$table += $obj
}

#Build a table for everyone who replied to the message
foreach ($result in $replytrack){
[string]$recipient = $result.Recipients
$obj = New-Object -TypeName PSObject
$obj | Add-Member -MemberType NoteProperty -Name Recipient -Value $recipient
$rftable += $obj
}

if ($rfcheck -eq "Y"){
#Add to the reply table for everyone who forwarded the message
    foreach ($result in $forwardtrack){
    [string]$recipient = $result.Recipients
    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name Recipient -Value $recipient
    $rftable += $obj
    }
}
#Remove duplicates from the list
$table = $table | select -Unique -Property recipient
$rftable = $rftable | select -Unique -Property recipient

#Find the message in all recipient mailboxes
foreach ($email in $table){
Write-Host Searching mailbox $email.Recipient
$MBsearch = Search-Mailbox -Identity $email.Recipient -TargetMailbox $targetMailbox -TargetFolder $targetFolderSearch -SearchDumpster -SearchQuery $searchQuery -WarningAction SilentlyContinue
[int]$itemCount = $MBsearch.ResultItemsCount
$resultsCount = $resultsCount + $itemCount
$resultsCountMB++
}

if ($rfcheck -eq "Y"){
#Find the forwarded or reply message
    foreach ($replyEmail in $rftable){
    $MBrfSearch = Search-Mailbox -Identity $replyEmail.Recipient -TargetMailbox $targetMailbox -TargetFolder $targetFolderSearch -SearchDumpster -SearchQuery $rfSearchQuery -WarningAction SilentlyContinue
    [int]$rfitemCount = $MBrfSearch.ResultItemsCount
    $RFResultsCount = $RFResultsCount + $rfitemCount
    }
}

#If there aren't any results, stop
if ($resultsCount -lt 1 -and $RFResultsCount -lt 1){
Write-Host Did not find any results -ForegroundColor Yellow
Break
}

Write-Host "Found $resultsCount messages in $resultsCountMB mailboxes, and $RFResultsCount replies/forwards" -ForegroundColor Green

$totalcount = $table.Count + $rftable.Count

#Found messages. Ask if they should be deleted
$title = "Message Deletion"
$message = "Please review the results in the mailbox $targetMailbox folder $targetFolderSearch. Continue with deletion?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
switch ($result)
    {
        0 {Write-Host Deleting message from $totalcount mailboxes
        foreach ($email in $table){
            Search-Mailbox -Identity $email.Recipient -TargetMailbox $targetMailbox -TargetFolder $targetFolderDelete -SearchQuery $searchQuery -DeleteContent -WarningAction SilentlyContinue -Confirm:$false
             }
        foreach ($rfemail in $rftable){
            Search-Mailbox -Identity $rfemail.Recipient -TargetMailbox $targetMailbox -TargetFolder $targetFolderSearch -SearchQuery $rfSearchQuery -DeleteContent -WarningAction SilentlyContinue -Confirm:$false
            }
        }
        1 {Break}
        default {Write-Host Please choose an option}
    }
