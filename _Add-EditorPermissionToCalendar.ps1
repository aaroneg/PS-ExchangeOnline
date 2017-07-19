Function Add-EditorPermissionToCalendar {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage='Which user is requesting access')][string[]]$RequestingUsers,
        [parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage='Which user will the access be applied to')][string[]]$TargetedUsers
    )

    Begin {
        Write-Verbose "Beginning Add-EditorPermissionToCalendar"
    }
    
    Process {
        Foreach ($RequestingUser in $RequestingUsers) {
            try {$RequestingUserObject=Get-ADUser $RequestingUser}
            catch {
                Write-Error "Could not find AD object for requesting user '$RequestingUser'"
                break
            }
            Write-Verbose "Processing requestor using object '$($RequestingUserObject.UserPrincipalName)'"
            Foreach ($TargetedUser in $TargetedUsers) {
                Write-Verbose "Current target $RequestingUser => $TargetedUser"
                try {$TargetedUserObject=Get-ADUSer $TargetedUser}
                catch {
                    Write-Error 'Targeted user $TargetedUser not found'
                    continue
                }
                Write-Verbose "Adding permissions for $($RequestingUserObject.UserPrincipalName) on target object $($TargetedUserObject.UserPrincipalName)"
                if (Get-MailboxFolderPermission -Identity "$($TargetedUserObject.UserPrincipalName):\Calendar"| Where-Object {$_.User.DisplayName -eq $RequestingUserObject.Name -and $_.AccessRights -ne 'Editor'}) {
                    Write-warning "Existing non-editor permissions found for $($RequestingUserObject.UserPrincipalName) on target $($TargetedUserObject.UserPrincipalName)"
                    Remove-MailboxFolderPermission -Identity "$($TargetedUserObject.UserPrincipalName):\Calendar" -User $RequestingUserObject.UserPrincipalName
                }
                
                if (Get-MailboxFolderPermission -Identity "$($TargetedUserObject.UserPrincipalName):\Calendar"| Where-Object {$_.User.DisplayName -eq $RequestingUserObject.Name -and $_.AccessRights -eq 'Editor'} ) {
                    Write-Verbose "Existing editor permissions found"
                    continue
                }

                Write-Verbose "Add-MailboxFolderPermission -Identity `"$($TargetedUserObject.UserPrincipalName):\Calendar`" -User $($RequestingUserObject.UserPrincipalName) -AccessRights Editor"
                Write-Verbose "Add-MailboxFolderPermission -Identity `"$($TargetedUserObject.UserPrincipalName):\Calendar`" -User $RequestingUser -AccessRights Editor"
                #region Shitty error handling
                # This is NOT the proper way to handle errors. Blame the shitty exchange cmdlets.
                Add-MailboxFolderPermission -Identity "$($TargetedUserObject.UserPrincipalName):\Calendar" -User $RequestingUserObject.UserPrincipalName -AccessRights Editor -ErrorAction SilentlyContinue
                Add-MailboxFolderPermission -Identity "$($TargetedUser):\Calendar" -User $RequestingUser -AccessRights Editor -ErrorAction SilentlyContinue
                #endregion
            }
        }
    }

    End {
        Write-Verbose "Ending Add-EditorPermissionToCalendar"
    }
}
