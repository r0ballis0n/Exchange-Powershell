#Initial User Message
Write-Host""
Write-Host "============================================================"
Write-Host "             Set Calendar Permisssion Script                "
Write-Host "                Attorneys & Investigators                   "
Write-Host "============================================================"
Write-Host ""
Write-Host "This script will go through the setup process for giving a  "
Write-Host "user access to another user's calendar. Please follow the    "
Write-Host "prompts to complete the process. Don't forget to contact the"
Write-Host "end user to verify that it worked as expected."
Write-Host ""

#Variables
$attorneymailboxes = @(get-ADGroupMember "Attorney Calendars" | ForEach-Object { Get-Mailbox $_.samAccountName })
$investigatormailboxes = @(Get-ADGroupMember "Investigator Calendars") | ForEach-Object { Get-Mailbox $_.samAccountName })

#User Prompt for Name
$User = Read-Host "Enter user's name to grant permissions to:"

#Commands
foreach ($attorneymailbox in $attorneymailboxes) {Add-MailboxFolderPermission -Identity $attorneymailbox":\Calendar" -User $User -AccessRights Editor}
foreach ($investigatormailbox in $investigatormailboxes ) {Add-MailboxFolderPermission -Identity $investigatormailbox":\Calendar" -User $User -AccessRights Editor}
