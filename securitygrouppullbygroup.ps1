#Iterates through a specified OU for AD groups, and then pulls users of the group.

Import-Module activedirectory

$folderPath = "C:\users\cfreeland\desktop\grouppull"
$check = test-path $folderPath
if (!$check) {
    mkdir $folderPath
}
#Pulls groups, and iterates through each group.
$groups = Get-ADGroup -filter * -SearchBase "OU=security,OU=Portland,DC=learning,DC=local" -Properties Name | `
% {

    $filepath = "C:\users\cfreeland\Desktop\GroupPull\"+$_.name+".csv" #specifies a local file path
       
    $members = Get-ADGroupMember -identity $_.name | Select-Object Name  #gets AD members for the current group that is being iterated through
    $_ | export-csv -Delimiter ',' -Path $filepath -NoTypeInformation -Force #exports group name
    $members| export-csv -Delimiter ',' -Path $filepath -NoTypeInformation -append -Force #exports members
}

