#Variables
Set-Variable -Name Author -Value "Author"
Set-Variable -Name Contributor -Value "Contributor"
Set-Variable -Name Editor -Value "Editor"
Set-Variable -Name None -Value "None"
Set-Variable -Name NonEditingAuthor -Value "NonEditingAuthor"
Set-Variable -Name Owner -Value "Owner"
Set-Variable -Name PublishingEditor -Value "PublishingEditor"
Set-Variable -Name PublishingAuthor -Value "PublishingAuthor"
Set-Variable -Name Reviewer -Value "Reviewer"

#Initial User Message
Write-Host""
Write-Host "============================================================"
Write-Host "             Set Calendar Permisssion Script                   "
Write-Host "============================================================"
Write-Host ""
Write-Host "This script will go through the setup process for modifing a  "
Write-Host "user's access to another user's calendar. Please follow the    "
Write-Host "prompts to complete the process. Don't forget to contact the"
Write-Host "end user to verify that it worked as expected."
Write-Host ""

#Functions
function Show-AccessRightsMenu
{
    param (
        [string]$Title = 'Access Rights Options'
    )
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' Editor."
    Write-Host "2: Press '2' Author."
    Write-Host "3: Press '3' Contributor."
    Write-Host "4: Press '4' Owner."
    Write-Host "5: Press '5' Reviewer."
    Write-Host "6: Press '6' NonEditing Author"
    Write-Host "7: Press '7' Publishing Editor"
    Write-Host "8: Press '8' Publishing Author"
    Write-Host "9: Press '9' None"
}

#User Input Prompts
$User = Read-Host -Prompt 'Enter requesting user email address "Example: user1@domain.com"'

$Identity = Read-Host -Prompt 'Enter the target user email address "Example: user2@domain.com"'

Show-AccessRightsMenu -Title 'Access Rights Options'
$AccessRightsOption = Read-Host "Please make a selection"
switch ($AccessRightsOption)
{
    1 {$AccessRights = $Editor}
    2 {$AccessRights = $Author}
    3 {$AccessRights = $Contributor}
    4 {$AccessRights = $Owner}
    5 {$AccessRights = $Reviewer}
    6 {$AccessRights = $NonEditingAuthor}
    7 {$AccessRights = $PublishingEditor}
    8 {$AccessRights = $PublishingAuthor}
    9 {$AccessRights = $None}
    
}

#Date
$Date = Get-Date

#Command
Set-MailboxFolderPermission -Identity $Identity":\Calendar" -User $User -AccessRights $AccessRights

#Final Output to user
Write-Host "You gave $User $AccessRights permissions to the calendar for $Identity on $Date. Please check with the end user that the process did complete as expected."
