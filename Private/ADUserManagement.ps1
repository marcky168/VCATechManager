# Function for AD User Management (Option 14)
function ADUserManagement {
    param([string]$AU, [pscredential]$Credential)

    Write-ConditionalLog "Starting AD User Management for AU $AU"

    $groupName = "H" + $AU.PadLeft(4, '0')  # e.g., 'H4048' for AU 4048

    try {
        # Query AD group members with credential and explicit server
        $groupMembers = Get-ADGroupMember -Identity $groupName -Server "vcaantech.com" -Credential $Credential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
        $users = $groupMembers | Get-ADUser -Properties Name, SamAccountName, EmailAddress, LockedOut, PasswordExpired, LastLogonDate -Server "vcaantech.com" -Credential $Credential -ErrorAction Stop

        if (-not $users) {
            Write-Host "No users found in group $groupName." -ForegroundColor Yellow
            Write-ConditionalLog "No users found in group $groupName."
            return
        }

        # Display users in a GridView for selection
        $selectedUser = $users | Select-Object Name, SamAccountName, EmailAddress, LockedOut, PasswordExpired, LastLogonDate | 
                        Out-GridView -Title "Select user for management in AU $AU" -OutputMode Single

        if ($selectedUser) {
            Write-Host "Selected user: $($selectedUser.Name) ($($selectedUser.SamAccountName))" -ForegroundColor Cyan
            $action = Read-Host "Choose action: (r)eset password, (u)nlock account, (c)ancel"

            switch ($action.ToLower()) {
                'r' {
                    $newPassword = Read-Host "Enter new password (will be converted to secure string)" -AsSecureString
                    if ($newPassword) {
                        Set-ADAccountPassword -Identity $selectedUser.SamAccountName -NewPassword $newPassword -Credential $Credential -ErrorAction Stop
                        Write-Host "Password reset successfully for $($selectedUser.SamAccountName)." -ForegroundColor Green
                        Write-ConditionalLog "Password reset for $($selectedUser.SamAccountName) in AU $AU."
                    } else {
                        Write-Host "No password entered. Cancelled." -ForegroundColor Yellow
                    }
                }
                'u' {
                    Unlock-ADAccount -Identity $selectedUser.SamAccountName -Credential $Credential -ErrorAction Stop
                    Write-Host "Account unlocked successfully for $($selectedUser.SamAccountName)." -ForegroundColor Green
                    Write-ConditionalLog "Account unlocked for $($selectedUser.SamAccountName) in AU $AU."
                }
                'c' {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                }
                default {
                    Write-Host "Invalid choice. Cancelled." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "No user selected." -ForegroundColor Yellow
            Write-ConditionalLog "No user selected for AU $AU."
        }
    } catch {
        Write-Host "Error fetching AD users for group $($groupName): $($_.Exception.Message)" -ForegroundColor Red
        Write-ConditionalLog "AD user fetch error for ${groupName}: $($_.Exception.Message)"
    }
}