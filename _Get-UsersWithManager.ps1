function Get-UsersWithManager {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage='Which manager? ex: user1')][string]$Manager,
        [Parameter(Mandatory=$falseHelpMessage='Include the manager in the list?')][switch]$IncludeManager
    )
    $Userlist=@()
    if ($IncludeManager) {$Userlist+=Get-ADUser -Properties manager,mail $Manager}
    $Userlist+=Get-ADUser -Properties manager,mail -Filter {manager -eq $Manager}|Where {$_.Enabled -eq $True}
    $Userlist|sort name
}
