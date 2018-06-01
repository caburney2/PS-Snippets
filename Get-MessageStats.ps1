
<#
 .DESCRIPTION
    Sends an email to the user with a table that contains the number of sent/received emails from their account
    over the last 7 day period.

 .NOTES
    1-Content is polled from Office 365 utilizing Get-MessageTrace
    2-365User requires an AD account with role:  Organization Management (or) Complianace Management (or) Help Desk in order to access
      Get-MessageTrace.  
 #>


#Office 365 connection credentials
$365User = "User.Name@contoso.com"
$File = 'C:\scripts\365creds.txt'
$365Credentials= New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $365User, (Get-Content $File | ConvertTo-SecureString)

[DateTime]$startDate = (Get-Date).AddDays(-7).ToString('MM/dd/yyy') + " 00:00 AM"
[DateTime]$endDate = "$(Get-Date -Format MM/dd/yyy) 11:59 PM"
$userList = Get-Content "C:\temp\userList.txt"

#SMTP Settings
$smtpFrom = "EmailStatistics@contoso.com"
$smtpServer = "SMTPRelay.contoso.com "

#Create Exchange Online Session.
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential Get-Credential -Authentication Basic -AllowRedirection

#Connect to Exchange Online session.
Import-PSSession -AllowClobber $Session

$userList | ForEach {
    $userEmailAddress = $_
    $userDisplayName = Get-ADUser -Filter {Mail -eq $userEmailAddress} -Properties Name | Select-Object -ExpandProperty Name
    $sentMessageCount = (Get-MessageTrace -SenderAddress $userEmailAddress -StartDate $startDate -EndDate $endDate).count
    $receivedMessageCount = (Get-MessageTrace -RecipientAddress $userEmailAddress -StartDate $startDate -EndDate $endDate).count

    #User Specific HTML Settings
    $smtpTo = $userEmailAddress
    $smtpSubject = "Email statistics for $userDisplayName"

    $html = " 
            <html>
            <head>
            <style>
                table {
                    font-family: arial, sans-serif;
                    border-collapse: collapse;
                    width: 50%;
                }
                
                td, th {
                    border: 1px solid #dddddd;
                    text-align: center;
                    padding: 8px;
                }
                
                tr:nth-child(even) {
                    background-color: #dddddd;
                }
            </style>
            </head>
                <body>
                <h4>$userDisplayName email statistics for $startDate-$endDate</h4>

                <table>
                    <tr>
                        <th>Sent Message Count</th>
                        <th>Received Message Count</th>
                    </tr>
                    <tr>
                        <td>$sentMessageCount</td>
                        <td>$receivedMessageCount</td>
                    </tr>
                 </table>

                 </body>
             </html>        

       "

       Send-MailMessage -SmtpServer $smtpServer -From $smtpFrom -To $userEmailAddress -Subject $smtpSubject -Body $html -BodyAsHtml
}