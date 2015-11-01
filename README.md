# OutlookCLI
A PowerShell CLI for MAPI


## Examples

```powershell
PS mail:\Inbox> $messages = Get-ChildItem
PS mail:\Inbox> $messages.Where({$_.SenderEmailAddress -eq "importantsencer@yourdomain.co"}) | Send-MailForward someemail@mydomain.me

PS mail:\Inbox> $messages = Get-ChildItem
PS mail:\Inbox> $messages.Where({$_.Subject -eq "Alert closed"}) | Set-MailRead

PS mail:\Inbox> $messages.Where({$_.Subject -eq "Alert Open"}) | Set-MailTask -Interval Today
```