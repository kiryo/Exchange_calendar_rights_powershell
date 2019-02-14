# Orginal code from Christopher Dargel
# https://www.bechtle-blog.com/home/bulk-set-calendar-permission-in-multilanguage-environments
# Set default as LimitedDetails for all calendars.
# Will get the language forch each users calendar Folder
# TEST WITH ONE ORE MORE USERS

# Edited /commented by kiryo 3.11.2018

# Go through all Mailboxes (add filter -like "John Doe" if you want to test), Add Domain Email to limit resultsize

foreach($mbx in Get-Mailbox -ResultSize Unlimited | where-object {$_.windowsemailaddress -like "*@yourdomain"})

{
# Get Name of user
$Calfolder = $Mbx.Name

# Add :\ to end of name
$Calfolder += ':\'

# Get word Calendar by right language!
$CalFolder += [string](Get-mailboxfolderstatistics $Mbx.alias -folderscope calendar | select -ExpandProperty name -first 1)

# Define variable
$mbx = $CalFolder

# Define test variable for error testing
$test = Get-MailboxFolderPermission -Identity $mbx -erroraction silentlycontinue

# Check if permissions are alredy done
$isright = Get-MailboxFolderPermission -Identity $mbx -user default | select -expandproperty AccessRights

# If permission is alredy defined and test came with no errors go and change permission
if ($isright -ne "LimitedDetails") {
if($test -ne $null)
{
# echo $mbx $isright  **UNCOMMENT THIS TO SHOW USERS WITH WRONG RIGHTS***
Set-MailboxFolderPermission -Identity $mbx -User Default -AccessRights LimitedDetails | out-null
}
}
}