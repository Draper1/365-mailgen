$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

$myPath = Get-Location
$myDir = $myPath.Path
$minlength = 1
$maxlength = 5
$minlines = 5
$maxlines = 1000
$maxrecipients = 10
$dictionary = "$($myDir)\dict.csv"
$textfile = "$($myDir)\longtext.txt"
#Add the 365 smtp server here e.g. draper365.mail.protection.outlook.com
$smtpServer = "draper365.mail.protection.outlook.com"
function EmailSubject {

	$subjectLength = Get-Random -Minimum $minlength -Maximum $maxlength
	[string]$subject = ""

	$i = 0
	do {

		$rand = Get-Random -Minimum 0 -Maximum ($wordcount - 1)
		$word = ($words.GetValue($rand)).Word
			
		$subject = $word + " " + $subject
		
		$i++
	}
	while ($i -lt $subjectLength)
	$subject = $subject.substring(0,1).ToUpper()+$subject.substring(1)
	
	return $subject
}

function EmailBody {
	$lineLength = Get-Random -Minimum $minlines -Maximum $maxlines
	[string]$body = ""

	$i = 0
	do {
		$rand = Get-Random -Minimum 0 -Maximum ($textlength - 1)
		$line = ($longtext.GetValue($rand))
		
		$body = $line + " " + $body
			
		$i++
	}
	while ($i -lt $lineLength)

	return $body
}

function PickAttachment {
    
    $rand = Get-Random -Minimum 0 -Maximum 10
    if ($rand -gt 7)
    {
        $attachfile = $true
    }
    else
    {
        $attachfile = $false
    }

    if ($attachfile)
    {
        $files = @(Get-ChildItem $myDir\Attachments | where { ! $_.PSIsContainer })
        $filecount = $files.Count
        $filepick = Get-Random -Minimum 0 -Maximum ($filecount)
        $file = $files.GetValue($filepick)
    }

    return $file
}

function PickRecipient {

	$rand = Get-Random -Minimum 0 -Maximum ($recipientcount)
	$name = $recipients.GetValue($rand)
	
	return $name
}


function PickSender {

	$rand = Get-Random -Minimum 0 -Maximum ($mailboxcount)
    $name = $mailboxes.GetValue($rand)
	
	return $name
}

function SendMail {

    Write-Host "*** New email message"

    #Generate subject, body and attachment
	$emailSubject = EmailSubject
	$emailBody = EmailBody
    $emailAttachment = PickAttachment

    #Choose sender for email
	$sender = PickSender
    
    $temp = Get-Mailbox -id $sender.id
	$SenderSmtpAddress = $temp.primarySmtpAddress
	$EmailSender = $SenderSmtpAddress
    Write-Host "Sender: $EmailSender"
    Write-Host "Subject: $EmailSubject"
    if ($emailAttachment)
    {
        Write-Host "Attachment: $emailAttachment"
    }

    $tocount = 1
    $i = 0
    do {
	    $recipient = PickRecipient
	    if ($recipient -eq $sender)
	    {
		    do { $recipient = PickRecipient }
		    while ($recipient -eq $sender)
	    }
        
        $tempObj = Get-Recipient -id $recipient.id
	    $RecipientSmtpAddress = $tempObj.PrimarySMTPAddress
	    $EmailRecipient = $RecipientSmtpAddress
        Write-Host "Recipient: $EmailRecipient"
        $i++
    }
    while ($i -lt $tocount)

    #Add the attachment
    if ($emailAttachment)
    {
        $attachmentPath = "$($myDir)\Attachments\$($emailAttachment)"

        $mailParam = @{
            To = $EmailRecipient
            From = $EmailSender
            Subject = $emailSubject
            Body = $emailBody
            SmtpServer = $smtpServer
            Port = "25"
            Credential = $UserCredential
            Attachments = $attachmentPath
        }
    } else {
        $mailParam = @{
            To = $EmailRecipient
            From = $EmailSender
            Subject = $emailSubject
            Body = $emailBody
            SmtpServer = $smtpServer
            Port = "25"
            Credential = $UserCredential
        }
    }
    Send-MailMessage @mailParam -UseSsl

}

if (Test-Path $dictionary)
{
	$words = @(Import-Csv $dictionary)
	$wordcount = $words.count
}
else
{
	Write-Host -ForegroundColor Yellow "Unable to locate dictionary file $dictionary."
	EXIT
}

if (Test-Path $textfile)
{
	$longtext = Get-Content $textfile
	$textlength = $longtext.count
}
else
{
	Write-Host -ForegroundColor Yellow "Unable to locate text file $textfile."
	EXIT
}


Write-Host -ForegroundColor White "Starting email generation loop"
do {
	[int]$hour = Get-Date -Format HH
	# You can modify these values to vary the number of emails that the
	# script will send each hour
	Switch($hour)
	{
		01 {$sendcount = 50}
		02 {$sendcount = 50}
		03 {$sendcount = 50}
		04 {$sendcount = 50}
		05 {$sendcount = 50}
		06 {$sendcount = 50}
		07 {$sendcount = 50}
		08 {$sendcount = 50}
		09 {$sendcount = 50}
		10 {$sendcount = 50}
		11 {$sendcount = 50}
		12 {$sendcount = 50}
		13 {$sendcount = 50}
		14 {$sendcount = 50}
		15 {$sendcount = 50}
		16 {$sendcount = 50}
		17 {$sendcount = 50}
		18 {$sendcount = 50}
		19 {$sendcount = 50}
		20 {$sendcount = 50}
		21 {$sendcount = 50}
		22 {$sendcount = 50}
		23 {$sendcount = 50}
		24 {$sendcount = 50}
		default {$sendcount = 50}
	}
	[string]$dayofweek = (Get-Date).Dayofweek
	Switch($dayofweek)
	{
		"Saturday"{$sendcount = 50}
		"Sunday" {$sendcount = 50}
    }
    
	Write-Host -ForegroundColor White "*** Will send $sendcount emails this hour"
    #Get list of mailbox users from 365
    $recipients = @()
    Write-Host -ForegroundColor White "Retrieving recipient list"
    $mailboxes = @(Get-Mailbox -RecipientTypeDetails UserMailbox -resultsize Unlimited | Where {$_.Name -ne "Administrator" -and $_.Name -notlike "extest_*"})
    $mailboxcount = $mailboxes.Count
    Write-Host "$mailboxcount mailboxes found"
    $recipients += $mailboxes
    $recipientcount = $recipients.count
    Write-Host "$recipientcount total recipients found"
    $sent = 0
    
	do {
		$pct = $sent/$sendcount * 100
		Write-Progress -Activity "Sending $sendcount emails" -Status "$sent of $sendcount" -PercentComplete $pct
        SendMail
        $sent++
        start-sleep -Seconds 30
	}
	until ($sent -eq $sendcount)
	Write-Host -ForegroundColor White "*** Finished sending $sendcount emails for hour $hour"
	
	# Check if there is any time still left in this hour
	# and sleep if there is
	[int]$endhour = Get-Date -Format HH
	if ($hour -lt 23)
	{
		[int]$nexthour = $hour + 1
		
			do {
				Write-Progress -Activity "Waiting for next hour to start" -Status "Sleeping..." -PercentComplete 0
				Write-Host -ForegroundColor Yellow "Not next hour yet, sleeping for 5 minutes"
				Start-Sleep 300
				[int]$endhour = Get-Date -Format HH
			}
			until($endhour -ge $nexthour)
		
	}
	else
	{
		[int]$nexthour = 0
		
			do {
				Write-Progress -Activity "Waiting for next hour to start" -Status "Sleeping..." -PercentComplete 0
				Write-Host -ForegroundColor Yellow "Not next hour yet, sleeping until hour $nexthour"
				Start-Sleep 300
				[int]$endhour = Get-Date -Format HH
			}
			until($endhour -eq $nexthour)
	}

}
until ($forever)
Remove-PSSession $Session