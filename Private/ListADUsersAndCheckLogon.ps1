# Function for List AD Users and Check Logon (Option 5)
function ListADUsersAndCheckLogon {
    param([string]$AU)

    Write-Log "Starting ListADUsersAndCheckLogon for AU $AU"

    try {
        # Ensure AD credentials are loaded
        if (-not $ADCredential) {
            Write-Host "AD credentials not loaded. Prompting for credentials..." -ForegroundColor Yellow
            $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)"
            if ($ADCredential) {
                $ADCredential | Export-Clixml -Path $credPathAD -Force
                Write-Host "AD credentials saved." -ForegroundColor Green
                Write-Log "AD credentials saved during function call"
            } else {
                Write-Host "No AD credentials provided. Cannot proceed." -ForegroundColor Red
                Write-Log "No AD credentials provided in function"
                return
            }
        }

        # Import PSTerminalServices module locally and verify
        $modulePath = "$global:ScriptRoot\Private\lib\PSTerminalServices"
        if (-not (Test-Path $modulePath)) {
            Write-Host "PSTerminalServices module path not found: $modulePath" -ForegroundColor Red
            Write-Log "PSTerminalServices module path not found"
            return
        }
        Import-Module -Name $modulePath -ErrorAction SilentlyContinue
        if (-not (Get-Command Get-TSSession -ErrorAction SilentlyContinue)) {
            Write-Host "PSTerminalServices module or Get-TSSession command not available." -ForegroundColor Red
            Write-Log "PSTerminalServices module not loaded"
            return
        }

        # Get users for the AU using Get-VCAHeadCount (similar to option 8 in VCAHospLauncher)
        $SiteUsers = Get-VCAHeadCount -AU $AU
        if (-not $SiteUsers.Users) {
            Write-Host "No AD users found for AU $AU." -ForegroundColor Yellow
            Write-Log "No AD users found for AU $AU"
            return
        }

        # Select user from the list
        $selectedUser = $SiteUsers.Users | Select-Object Name, SamAccountName, Title, Department | Sort-Object Name | Out-GridView -Title "Select AD User for AU $AU (Total: $($SiteUsers.UserCount))" -OutputMode Single
        if (-not $selectedUser) {
            $username = Read-Host "No user selected. Enter username to check logon"
            if (-not $username) {
                Write-Host "No username provided." -ForegroundColor Yellow
                return
            }
        } else {
            $username = $selectedUser.SamAccountName
        }
        Write-Log "Selected user: $username"

        # Display AD properties for the selected user
        try {
            # Get domain password policy for expiry calculation
            $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -Server "vcaantech.com" -Credential $ADCredential).MaxPasswordAge
            $adUser = Get-ADUser -Identity $username -Properties Name, Title, OfficePhone, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop
            Write-Host "`nAD Properties for $username :" -ForegroundColor Cyan
            $adUser | Select-Object Name, Title, @{n='OfficePhone'; e={$_.OfficePhone}}, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet, @{n='PasswordExpires'; e={ if ($_.PasswordLastSet) { $_.PasswordLastSet + $MaxPasswordAge } else { 'Never Set' } }} | Format-List
        } catch {
            Write-Host "User '$username' not found in AD. Proceed anyway? (y/n)" -ForegroundColor Yellow
            if ((Read-Host).ToLower() -ne 'y') { return }
        }

        # Sequential Get-TSSession check for active sessions (run locally with -ComputerName, no credentials)
        $servers = Get-CachedServers -AU $AU

        $activeSessions = @()
        $totalServers = $servers.Count
        $i = 0
        foreach ($server in $servers) {
            $i++
            Write-Progress -Activity "Checking active sessions for $username" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
            try {
                $sessions = Get-TSSession -ComputerName $server -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Active' } | Where-Object { 
                    # Improved filter to handle domain prefixes (e.g., VCAANTECH\Eun.An)
                    $sessionUser = $_.UserName
                    $sessionUser -eq $Username -or $sessionUser -like "*\$Username" -or $sessionUser -like "*\\$Username" -or $sessionUser -eq "VCAANTECH\$Username"
                }
                foreach ($session in $sessions) {
                    $activeSessions += [PSCustomObject]@{
                        Server     = $server
                        UserName   = $session.UserName
                        SessionId  = $session.SessionId
                        State      = $session.State
                        LogOnTime  = $session.LogOnTime
                        ClientIP   = $session.IPAddress  # Use direct IP if available
                        ClientName = $session.ClientName  # Add ClientName for DNS resolution
                    }
                }
            } catch {
                # Error handled silently
            }
        }

        Write-Progress -Activity "Checking active sessions for $username" -Completed
        Write-Log "Active sessions check completed: Found $($activeSessions.Count) sessions"

        # Resolve IP addresses via DNS if ClientIP is empty, preferring IPv4
        foreach ($session in $activeSessions) {
            if (-not $session.ClientIP -or $session.ClientIP -eq "N/A" -or $session.ClientIP -eq "") {
                if ($session.ClientName -and $session.ClientName -ne "") {
                    try {
                        $addresses = [System.Net.Dns]::GetHostAddresses($session.ClientName)
                        # Prefer IPv4 over IPv6
                        $resolvedIP = $addresses | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 | Select-Object -ExpandProperty IPAddressToString
                        if (-not $resolvedIP) {
                            $resolvedIP = $addresses | Select-Object -First 1 | Select-Object -ExpandProperty IPAddressToString
                        }
                        $session.ClientIP = $resolvedIP
                    } catch {
                        $session.ClientIP = "N/A"
                    }
                } else {
                    $session.ClientIP = "N/A"
                }
            }
        }

        if ($activeSessions) {
            $selectedSession = $activeSessions | Out-GridView -Title "Active Sessions for $username on AU $AU - Select to Launch VNC/Shadow" -OutputMode Single
            if ($selectedSession) {
                $launchChoice = Read-Host "Launch VNC (v) or RDP Shadow (r) for $($selectedSession.UserName) on $($selectedSession.Server)? (v/r/n)"
                if ($launchChoice.ToLower() -eq 'v') {
                    $vncPath = "$global:ScriptRoot\Private\bin\vncviewer.exe"
                    if (Test-Path $vncPath) {
                        $userIP = $selectedSession.ClientIP
                        if ($userIP -and $userIP -ne "N/A" -and $userIP -ne "") {
                            try {
                                # Validate executable before launching
                                $fileInfo = Get-Item $vncPath -ErrorAction Stop
                                if ($fileInfo.Length -lt 1000) {
                                    throw "VNC executable appears to be corrupted or incomplete (file size: $($fileInfo.Length) bytes)"
                                }
                                Start-Process $vncPath -ArgumentList $userIP -ErrorAction Stop
                                Write-Host "Launching VNC for $($selectedSession.UserName) on $userIP." -ForegroundColor Green
                                Write-Log "Launched VNC for $($selectedSession.UserName) on $userIP"
                            } catch {
                                $errorMessage = $_.Exception.Message
                                Write-Host "Failed to launch VNC viewer: $errorMessage" -ForegroundColor Red
                                Write-Log "VNC launch failed: $errorMessage"
                                if ($errorMessage -like "*not a valid application*") {
                                    Write-Host "The VNC executable appears to be corrupted or incompatible. Please re-download it." -ForegroundColor Yellow
                                }
                            }
                        } else {
                            Write-Host "No IP address available for VNC." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "VNC viewer not found at $vncPath." -ForegroundColor Yellow
                    }
                } elseif ($launchChoice.ToLower() -eq 'r') {
                    Start-Process "mstsc.exe" -ArgumentList @("/v:$($selectedSession.Server)", "/shadow:$($selectedSession.SessionId)", "/control")
                    Write-Host "Launching RDP Shadow for $($selectedSession.UserName) on $($selectedSession.Server)." -ForegroundColor Green
                    Write-Log "Launched RDP Shadow for $($selectedSession.UserName) on $($selectedSession.Server)"
                } else {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "No active sessions found for $username. Falling back to logon event search." -ForegroundColor Yellow
            Write-Log "No active sessions for $username, calling User-LogonCheck"
            User-LogonCheck -AU $AU -Username $username -SkipPropertiesDisplay
        }
    } catch {
        Write-Host "Error in ListADUsersAndCheckLogon: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error in ListADUsersAndCheckLogon: $($_.Exception.Message)"
    }
}