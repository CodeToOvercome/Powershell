# Useful Powershell cmdlets

## Exchange related

To get inbox rules of a mailbox

    Get-inboxrule -mailbox username | Select-Object Name, Description | Format-List

To set auto reply for a user

    Set-MailboxAutoReplyConfiguration -Identity username -AutoReplyState Enabled -InternalMessage "Internal auto-reply message." -ExternalMessage "External auto-reply message."

To list members of mailbox with permissions

    Get-MailboxPermission -Identity mailboxName@domain.com |Select-Object User,AccessRights,IsInherited |Where-Object{$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $False}

To remove a user from  shared mailbox

    Remove-MailboxPermission -Identity mailboxName@domain.com -user username -AccessRights FullAccess

To get list of mailboxes a user has access to

    Get-Mailbox |Get-MailboxPermission -User username

To copy mailbox data from one to another

    Search-Mailbox -Identity sourceMailbox@domain.com -TargetMailbox destMailboxName -TargetFolder "DestMailboxFolder"

To get list of all shared mailbox

    Get-Mailbox -Filter * -RecipientTypeDetails Shared  |Select-Object Name,PrimarySmtpAddress,RecipientTypeDetails,AccountDisabled |Export-Csv -Path 'filepath\filename.csv'

To get list of members in Dynamic Distribution groups

 >First save the group in a variable

    $FTE = Get-DynamicDistributionGroup "fancyGroupName"

>Then run the following command

    Get-Recipient -RecipientPreviewFilter $FTE.RecipientFilter -OrganizationalUnit $FTE.RecipientContainer |Sort-Object Name

To get SendOnBehalfto permissions from a shared mailbox

    Get-Mailbox -Identity mailboxName |Select-Object GrantSendonBehalfto

To get to know what type of mailbox is something is

    Get-Mailbox -Identity $user | Select-Object RecipientTypeDetails

To get Calendar Permissions level on a mailbox

    Get-MailboxFolderPermission -Identity username@domain.com:\calendar

### Mailbox Forwarding

To get information on mailbox forwarding

    Get-Mailbox username |Format-List delivertomailboxandforward,forwardingaddress

To set mailbox forwarding with a local copy enabled

    Set-Mailbox username -ForwardingAddress username@domain.com

To remove mail forwarding

    Set-Mailbox Bradp@o365info.com -ForwardingAddress $Null

### State of Mailbox

To find inactive mailbox in your exchange

    Get-MessageTrace -RecipientAddress username@domain.com -StartDate 06/06/2019 -EndDate 07/05/2019

To get list of all active mailboxes

    Get-Mailbox -Filter * -RecipientTypeDetails UserMailbox |Where-Object {$_.AccountDisabled -eq $false }|Format-List -Property PrimarySmtpaddress |export-csv -Path 'folderPath\fileName.csv'

### Mailbox Statistics

To get mailbox folder statistics

    Get-MailboxFolderStatistics -Identity username -Archive |Select-Object Name,FolderSize |Sort-Object FolderSize

To get mailbox size

    Get-MailboxStatistics -Identity username | Select-Object DisplayName, @{n=”Total Size (MB)”;e={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, StorageLimitStatus

## AD related cmdlets

To get list of users in a site or container

    Get-ADUser -Filter * -SearchBase "OU=Name,OU=Sites,OU=name,DC=name,DC=internal"

To get users with telephone number in a container

    Get-ADUser -Filter * -SearchBase "OU=Name,OU=Sites,OU=name,DC=name,DC=internal" -Properties TelephoneNumber,DisplayName | Select-Object DisplayName,TelephoneNumber

To get list of disabled users in the whole Org

    Get-ADUser -Filter * -SearchBase "OU=name,DC=name,DC=internal" -Properties DisplayName,Enabled | Select-Object DisplayName,Enabled |Where-Object{$_.Enabled -eq $false}

To rename a user's CanonicalName

    Rename-ADObject -Identity "9bf48c17-dfai-4b5a-8c9c-5b1afdba21c5" -NewName "newName"

To rename a user's DisplayName

    Set-ADUser -Identity username -DisplayName "newName"

To list all distribution groups a user belongs to

    Get-ADPrincipalGroupMembership -Identity username |Select-Object Name,GroupCategory|Where-Object{$_.GroupCategory -eq "Distribution"}

To get Distribution groups in local AD with e-mail address

    Get-ADGroup -Filter 'GroupCategory -eq "Distribution"' -Properties Mail |Select-Object Name,Mail

To get list of members in distribution group including the nested groups

    Get-ADGroupMember groupName -Recursive  | Select-Object Name |Sort-Object Name

To list all currently disabled users in AD

    Get-ADUser -Filter * |Select-Object Name,Enabled |Where-Object{$_.Enabled -eq $false}| Sort-Object Name

## Misc

To get to know the media type of physical disk

    Get-PhysicalDisk |Select-Object MediaType

To find serial number of the laptop hardware

    Get-WmiObject win32_bios |Format-List serialnumber

To enter remote session using Powershell

    Enter-PSSession -ComputerName mycomputerName

To get permission settings on file share or a folder

    Get-ChildItem .\folderName -Recurse |Where-Object{($_.PsIsContainer)}|Get-Acl |Format-List -Property Path,Owner,Group,AccessToString | Out-File 'folderPath/filename.csv'
