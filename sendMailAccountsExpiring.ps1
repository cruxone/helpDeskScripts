Import-Module activedirectory
$emailSmtpServer = "smtp.office365.com"
$emailSmtpServerPort = "587"
$emailSmtpUser = Read-Host -AsSecureString
$emailSmtpPass = Read-Host -AsSecureString
 
$expiringAccounts = Search-ADAccount -SearchBase "OU=Portland,DC=learning,DC=local" -AccountExpiring -TimeSpan 45.00:00:00 | `
                    where {$_.ObjectClass -eq 'user'}
 
$exportUsers = @()
                   
ForEach ($user in $expiringAccounts) {
    $exportUser = get-aduser -Identity $user -Properties AccountExpirationDate,Manager,Enabled |`
    select AccountExpirationDate,name,Enabled,@{N='Manager';E={(Get-ADUser $_.Manager).samaccountName}} | Sort-Object Manager
    $exportUsers += $exportUser
    
    $userManager = (get-aduser (get-aduser -identity $user -Properties manager).manager).samaccountname

    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = $emailSmtpUser
    $emailMessage.To.Add( $userManager + "@learning.com")
    $emailMessage.Subject = $user.Name + " Account is expiring"
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = "Hello,<br /> <br />You are receiveing this email because you are the registered manager for <b>" + $user.Name `
                         + "</b>. The account for <b>"  + $user.Name + "</b> is going to expire on <b>" `
                         + $user.AccountExpirationDate `
                         + "</b>. We can extend this by up to one year. If you would like an extension to this account " `
                         + "or the manager of the account to be changed, please email helpdesk@learning.com." `
                         + "<br /> <br />Thank you,<br />CF"
 
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($emailSmtpServer,$emailSmtpServerPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($emailSmtpUser,$emailSmtpPass);
    $SMTPClient.Send($emailMessage)
}
 
$exportUsers | Export-Csv -Delimiter "," C:\users\cfreeland\Desktop\expiringAccounts.csv -NoTypeInformation
write $exportUsers
