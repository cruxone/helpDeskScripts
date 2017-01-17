Import-Module activedirectory

$emailSmtpServer = "smtp.office365.com"
$emailSmtpServerPort = "587"
write-host "[#] Enter username/email address."
$emailSmtpUser = Read-Host
write-host "[#] Enter password."
$emailSmtpPass = Read-Host -AsSecureString
write-host "[#] Enter email address of IT Department."
$itDepartmentEmail = read-host
$SMTPClient = New-Object System.Net.Mail.SmtpClient($emailSmtpServer,$emailSmtpServerPort)
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($emailSmtpUser,$emailSmtpPass);
function emailTeam ($file) {
    $teamMessage = New-Object System.Net.Mail.MailMessage
    $teamMessage.From = $emailSmtpUser
    $teamMessage.To.Add($itDepartmentEmail)
    $teamMessage.Subject = "Active Directory Accounts expiring"
    $teamMessage.IsBodyHtml = $true
    $teamMessage.Body = "Hello,<br /> <br />You are receiveing this email from a monthly check for expiring AD accounts. " `
                         + "Attached is a csv containing expiring accounts." `
                         + "<br /> <br />Thank you,<br />CF"
    $attach = new-object Net.Mail.Attachment($file)
    $teamMessage.Attachments.add($attach)
    $SMTPClient.Send($teamMessage)
}
function emailManagers {
    $filePath = "C:\users\cfreeland\Desktop\expiringAccounts.csv"
    $expiringAccounts = Search-ADAccount -SearchBase "OU=Portland,DC=learning,DC=local" -AccountExpiring -TimeSpan 45.00:00:00 | `
                    where {$_.ObjectClass -eq 'user'}
    $exportUsers = @()
    ForEach ($user in $expiringAccounts) {
        $exportUser = get-aduser -Identity $user -Properties AccountExpirationDate,Manager,Enabled | `
        select AccountExpirationDate,name,Enabled,@{N='Manager';E={(Get-ADUser $_.Manager).samaccountName}} | Sort-Object Manager
        $exportUsers += $exportUser
        $userManager = (get-aduser (get-aduser -identity $user -Properties manager).manager).samaccountname
        $managerMessage = New-Object System.Net.Mail.MailMessage
        $managerMessage.From = $emailSmtpUser
        $emailMessage.To.Add( $userManager + "@learning.com")
        $managerMessage.Subject = $user.Name + " Account is expiring"
        $managerMessage.IsBodyHtml = $true
        $managerMessage.Body = "Hello,<br /> <br />You are receiveing this email because you are the registered manager for <b>" + $user.Name `
                             + "</b>. The account for <b>"  + $user.Name + "</b> is going to expire on <b>" `
                             + $user.AccountExpirationDate `
                             + "</b>. We can extend this by up to one year. If you would like an extension to this account " `
                             + "or the manager of the account to be changed, please email helpdesk@learning.com." `
                             + "<br /> <br />Thank you,<br />CF"
        $SMTPClient.Send($managerMessage)
    }
    $exportUsers | Export-Csv -Delimiter "," $filePath -NoTypeInformation
    write $exportUsers
    emailTeam($filePath)
}
emailManagers
