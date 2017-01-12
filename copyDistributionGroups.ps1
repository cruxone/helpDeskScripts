#####################################################
##                                                 ##
##   Script to grab a users distribution groups,   ##
##   Then add a different user to those groups     ##
##   Chris Freeland 3/18/2016                      ##
##                                                 ##
#####################################################


### Creates O365 session within current PoSh window ###
$usercredential = get-credential
$session = new-pssession -configurationname microsoft.exchange -connectionuri https://outlook.office.com/powershell-liveid/ -credential $usercredential -authentication basic -allowredirection
import-pssession $session

$userToCopy = Read-Host -Prompt 'User name to copy: ' 
$userToAdd = Read-Host -Prompt 'User name to add: '

Write-Host 'retreiving distribution groups associated with user'

$Mailbox = get-Mailbox $userToCopy 
$DN = $mailbox.DistinguishedName  ##Pulls distinguished name from mailbox
$Filter = "Members -like ""$DN""" ##Extracts users with a matching distinguished name. 

Get-DistributionGroup -ResultSize Unlimited -Filter $Filter | select displayname | export-CSV C:\users\cfreeland\desktop\groupsToAdd.csv -NoTypeInformation

$usergroups = import-csv C:\users\cfreeland\desktop\groupsToAdd.csv 

##Adds user to groups and exports a csv with groups added to.
ForEach ($usergroup in $usergroups) {
    try {
        add-DistributionGroupMember -identity $usergroup.displayname -Member $userToAdd
        Write-Host 'Added' $userToAdd 'to' $usergroup
    } catch {
        write-host 'an error has occured'
    }
}
