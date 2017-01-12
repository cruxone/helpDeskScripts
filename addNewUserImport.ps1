# Pulls in a CSV, creates user with the specified information, then writes out accounts created
# CSV requires these fields: Firstname,Lastname,Maildomain,OU,Password

Import-Module ActiveDirectory

$Users = Import-Csv -Path "C:\Users\cfreeland\Documents\WorkDocs\userlist-sn.csv"            
foreach ($User in $Users) {            
    if (!$User.Firstname) {
        Write-Host "EOD"
        return
    }

    $Displayname = $User.Firstname + " " + $User.Lastname
    $name = $Displayname         
    $UserFirstname = $User.Firstname            
    $UserLastname = $User.Lastname            
    $OU = $User.OU     
    $SAM = $User.Firstname[0] + $User.Lastname            
    $UPN = $User.Firstname[0] + $User.Lastname + "@" + $User.Maildomain            
    $Description = "Extranet account for " + $Displayname            
    $Password = $User.Password
    
    New-ADUser -Name $name -DisplayName $Displayname -SamAccountName $SAM `
    -UserPrincipalName $UPN -GivenName $UserFirstname -Surname $UserLastname -Description $Description `
    -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true `
    -Path $OU -ChangePasswordAtLogon $false –PasswordNeverExpires $true          
    
    write-host "Account created for" $name"," $UPN
}
