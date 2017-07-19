if (!($ExOnlineSession)) {
    if (!$o365Cred) {
        $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
        $email=$searcher.FindOne().Properties.mail
        $o365Cred=Get-Credential -Message "o365 Admin" -UserName $email
    }
    $ExOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365cred -Authentication Basic -AllowRedirection
    Import-PSSession $ExOnlineSession
}
function Remove-ExchangeOnlineConnection {
    Remove-PSSession $ExOnlineSession
}
