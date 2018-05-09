<# 
This depends on you having every employee's manager properly filled out in AD. If you don't have
that, this script will do nothing for you. Also, this script only defines the functions, it does
not run them.
#>
. $PSScriptRoot\_initialize-OnlineExchange.ps1
. $PSScriptRoot\_Add-EditorPermissionToCalendar.ps1
. $PSScriptRoot\_Get-UsersWithManager.ps1

Function Assign-AccessForTeamToPerson {
    <# 
        .SYNOPSIS
        Adds access for a team of people to 1 person's calendar
        .DESCRIPTION
        Provide a target user and manager to allow access for an entire team to that person's calendar.
    #>
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,HelpMessage="Which person manages the team that will need access to this calendar? Example: Manager1")][string]$Manager,
        [parameter(Mandatory=$true,HelpMessage="Which person's calendar are we allowing the team access TO? Example: User1")][string]$TargetUser
    )
    Try {Get-ADUser $Manager|out-null}
    Catch {Throw "User '$Manager' not found."}
    Try {Get-Aduser $TargetUser|out-null}
    Catch {Throw "User '$TargetUser' not found."}
    $RequestingUsers=(Get-UsersWithManager -Manager $Manager -IncludeManager).SamAccountName
    Add-EditorPermissionToCalendar -RequestingUsers $RequestingUsers -TargetedUsers $TargetUser 
}
Function Assign-AccessForPersonToTeam {
    <# 
        .SYNOPSIS
        Adds access for 1 person to a team's calendar
        .DESCRIPTION
        Provide a requesting user and manager to allow access for a person to an entire team's calendar.
    #>
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,HelpMessage="Which person manages the team that will need access to this calendar? Example: Manager1")][string]$Manager,
        [parameter(Mandatory=$true,HelpMessage="Which person's calendar are we allowing the team access TO? Example: User1")][string]$RequestingUser
    )
    Try {Get-ADUser $Manager|out-null}
    Catch {Throw "User '$Manager' not found."}
    Try {Get-Aduser $RequestingUser|out-null}
    Catch {Throw "User '$RequestingUser' not found."}
    $TargetUsers=(Get-UsersWithManager -Manager $Manager -IncludeManager).SamAccountName
    $TargetUsers
    Add-EditorPermissionToCalendar -RequestingUsers $RequestingUser -TargetedUsers $TargetUsers
}
