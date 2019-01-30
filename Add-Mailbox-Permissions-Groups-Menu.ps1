#Initial User Message
Write-Host""
Write-Host "============================================================"
Write-Host "             Add Calendar Permisssion Script                "
Write-Host "                Specified AD Group/Groups                   "
Write-Host "============================================================"
Write-Host ""
Write-Host "This script will go through the setup process for giving a  "
Write-Host "user access to members of an ad secruity group's calendar.  "
Write-Host "Please follow the prompts to complete the process. Don't "
Write-Host "forget to contact the end user to verify that it worked as "
Write-Host "expected."
Write-Host ""

#Variables
$adgroup1 = @(get-ADGroupMember "AD Group 1" | ForEach-Object { Get-Mailbox $_.samAccountName })
$adgroup2 = @(Get-ADGroupMember "AD Group 2") | ForEach-Object { Get-Mailbox $_.samAccountName })

#User Prompt for Name
$User = Read-Host "Enter user's name to grant permissions to"

#Commands
foreach ($adgroup1mailbox in $adgroup2mailboxes) {Add-MailboxFolderPermission -Identity $adgroup1mailbox":\Calendar" -User $User -AccessRights Editor}
foreach ($adgroup2mailbox in $adgroup2mailboxes ) {Add-MailboxFolderPermission -Identity $adgroup2mailbox":\Calendar" -User $User -AccessRights Editor}
