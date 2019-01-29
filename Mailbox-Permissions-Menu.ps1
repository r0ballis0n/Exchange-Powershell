Write-Host""
Write-Host "============================================================"
Write-Host "               Mailbox Permisssion Script                   "
Write-Host "============================================================"
Write-Host ""
Write-Host "This script will go through the setup process for giving a  "
Write-Host "user access to another user's mailbox. Please follow the    "
Write-Host "prompts to complete the process. Don't forget to contact the"
Write-Host "user to verify that it worked as expected."
Write-Host ""

#Variables
Set-Variable -Name FullAccess -Value "FullAccess"
Set-Variable -Name ReadPermission -Value "ReadPermission"
Set-Variable -Name ChangeOwner -Value "ChanageOwner"
Set-Variable -Name ChangePermission -Value "ChangePermission"
Set-Variable -Name DeleteItem -Value "DeleteItem"
Set-Variable -Name ExternalAccount -Value "ExternalAccount"

#Functions
function Show-PermissionMenu
{
    param (
        [string]$Title = 'Permission Options'
    )
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' Full Access."
    Write-Host "2: Press '2' Read Permission."
    Write-Host "3: Press '3' Change Owner."
    Write-Host "4: Press '4' Change Permission."
    Write-Host "5: Press '5' Delete Item."
    Write-Host "6: Press '6' External Account"
}

function Show-AutoMapMenu
{
    param (
        [string]$Title = 'Outlook Auto Mapping'
    )
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' Outlook Auto Mapping Enabled."
    Write-Host "2: Press '2' Outlook Auto Mapping Disabled."
}

#User Inputs
$User = Read-Host -Prompt 'Enter requesting user name "Example: Mary User"'

$Identity = Read-Host -Prompt 'Enter the mailbox name "Example: Todd User"'

Show-PermissionMenu -Title 'Permission Options'
$AccessRightsOption = Read-Host "Please make a selection"
switch ($AccessRightsOption)
{
    1 {$AccessRights = $FullAccess}
    2 {$AccessRights = $ReadPermission}
    3 {$AccessRights = $ChangeOwner}
    4 {$AccessRights = $ChangePermission}
    5 {$AccessRights = $DeleteItem}
    6 {$AccessRights = $ExternalAccount}
    
}

Show-AutoMapMenu -Title 'Outlook Auto Mapping'
$AutoMapping = Read-Host "Please make a selection"
switch ($AutoMapping)
{
    1 {$AutoMappingConverted = $true; break}
    2 {$AutoMappingConverted = $false; break}
}

$Date = Get-Date

#Command
Add-MailboxPermission -Identity  "$Identity" -User "$User" -AccessRights $AccessRights -AutoMapping $AutoMappingConverted

#Final Output to user
Write-Host "You gave $User $AccessRights to the mailbox for $Identity on $Date. Please check with the user that the process did complete as expected."
