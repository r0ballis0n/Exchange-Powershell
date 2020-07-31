$mailbox = Read-Host "What Mailbox do you want to give access to?"
$username = Read-Host "Who needs the access?"

Add-MailboxPermission -identity $mailbox -User $username -AccessRights FullAccess -Automapping:$false