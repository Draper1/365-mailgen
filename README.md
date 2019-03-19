# Office 365 Mailgen

## Usage

In order to use this script, first determine the 365 protection MX domain, and update the variable $smtpServer e.g.

```
$smtpServer = "draper365.mail.protection.outlook.com"
```

Save the script and ensure all 3 files exist in the directory:

* longtext.txt
* dict.csv
* 365_mail_gen.ps1

Run the Powershell script and confirm no errors are being seen from sending emails.

Please be aware, using this places like AWS will put SMTP restrictions in place when using port 25, update the `@mailparam` values if you'd like to use different mail settings with Send-MailMessage.

If you would like to increase the number of mails being sent, refer to `Write-Host -ForegroundColor White "Starting email generation loop"`
You can specify the number of mails by the hour.

We will also sleep the script after each mail for 30 seconds, this can be increased or reduced:

Find `start-sleep -Seconds 30` and change the value in seconds.