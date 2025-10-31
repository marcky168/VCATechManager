# Consolidated Network Functions
# Contains: DeviceConnectivityTest

function DeviceConnectivityTest {
    param([string]$AU)

    Write-Host "Searching for Heska devices in DHCP leases and reservations for AU $AU" -ForegroundColor Cyan
    Write-ConditionalLog "Starting Heska Device Search for AU $AU"

    # Load admin credentials for DHCP server access (required for permissions)
    $adminCredPath = "$global:ScriptRoot\Private\vcaadmin.xml"
    if (Test-Path $adminCredPath) {
        try {
            $AdminCredential = Import-Clixml -Path $adminCredPath
            Write-ConditionalLog "Admin credentials loaded for DHCP access in AU $AU"
        } catch {
            Write-Host "Failed to load admin credentials from $adminCredPath : $($_.Exception.Message). DHCP access may fail." -ForegroundColor Red
            Write-ConditionalLog "Failed to load admin credentials for DHCP: $($_.Exception.Message)"
            $AdminCredential = $null
        }
    } else {
        Write-Host "Admin credentials file not found at $adminCredPath. Update via menu option 11. DHCP access may fail." -ForegroundColor Yellow
        Write-ConditionalLog "Admin credentials file missing for DHCP access in AU $AU"
        $AdminCredential = $null
    }

    # Get DHCP servers for this AU
    $dhcpServers = Get-DHCPServersForAU -AU $AU -Credential $ADCredential
    if (-not $dhcpServers -or $dhcpServers.Count -eq 0) {
        Write-Host "No DHCP servers configured or discovered for AU $AU. Skipping search." -ForegroundColor Yellow
        return
    }

    # Get scopes from all DHCP servers, filtered for this AU (like VCAHospLauncher option 24b)
    Write-Host "Retrieving DHCP scopes from all servers for AU $AU..." -ForegroundColor Cyan
    $allScopes = @()
    foreach ($server in $dhcpServers) {
        try {
            $serverScopes = Invoke-Command -ComputerName $server -ScriptBlock { param($au) Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$au-*" } } -ArgumentList $AU -Credential $AdminCredential -ErrorAction Stop
            if ($serverScopes) {
                # Add server info to each scope object
                foreach ($scope in $serverScopes) {
                    $scope | Add-Member -MemberType NoteProperty -Name "DHCPServer" -Value $server -Force
                    $allScopes += $scope
                }
            }
            Write-Host "Found $($serverScopes.Count) scopes on $server." -ForegroundColor Green
        } catch {
            Write-Host "Error fetching scopes from $server : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($allScopes.Count -eq 0) {
        Write-Host "No DHCP scopes found on any DHCP server for AU $AU." -ForegroundColor Red
        return
    }

    Write-Host "Found $($allScopes.Count) total scopes across all DHCP servers for AU $AU." -ForegroundColor Green

    # Select a scope from all available scopes (deduplicate by ScopeId)
    if ($allScopes.Count -eq 1) {
        $selectedScope = $allScopes[0]
        Write-Host "Using DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    } else {
        # Deduplicate scopes by ScopeId, keeping the first occurrence (first server found)
        $uniqueScopes = $allScopes | Group-Object -Property ScopeId | ForEach-Object {
            $_.Group | Select-Object -First 1
        }
        $scopeDisplay = $uniqueScopes | Select-Object Name, ScopeId, StartRange, EndRange, Description, LeaseDuration, SubnetMask, DHCPServer
        $selectedScope = $scopeDisplay | Out-GridView -Title "Select DHCP Scope for AU $AU" -OutputMode Single
        if (-not $selectedScope) {
            Write-Host "No DHCP scope selected. Operation cancelled." -ForegroundColor Yellow
            return
        }
        # Get the full scope object back from the original scopes array
        $selectedScope = $allScopes | Where-Object { $_.ScopeId -eq $selectedScope.ScopeId } | Select-Object -First 1
        Write-Host "Selected DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    }

    $scopeId = $selectedScope.ScopeId
    $selectedServer = $selectedScope.DHCPServer
    Write-Host "Selected scope: $scopeId ($($selectedScope.Name)) on server $selectedServer" -ForegroundColor Green

    $heskaDevices = @()

    # Search the selected DHCP server and scope
    Write-Host "Searching DHCP server: $selectedServer for scope: $scopeId" -ForegroundColor Cyan

    # Get DHCP leases
    try {
        $leases = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Lease -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($leases.Count) DHCP leases from $selectedServer." -ForegroundColor Cyan

        # Filter for Heska devices in leases
        $heskaLeases = $leases | Where-Object {
            ($_.HostName -and $_.HostName -ilike "*heska*") -or
            ($_.Description -and $_.Description -ilike "*heska*")
        }

        if ($heskaLeases) {
            foreach ($lease in $heskaLeases) {
                $heskaDevices += [PSCustomObject]@{
                    Type        = "Lease"
                    IPAddress   = $lease.IPAddress
                    ClientId    = $lease.ClientId
                    HostName    = $lease.HostName
                    Description = $lease.Description
                    LeaseExpiry = $lease.LeaseExpiryTime
                }
            }
        }
    } catch {
        Write-Host "Error fetching leases from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Heska search"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Get DHCP reservations
    try {
        $reservations = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Reservation -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($reservations.Count) DHCP reservations from $selectedServer." -ForegroundColor Cyan

        # Filter for Heska devices in reservations
        $heskaReservations = $reservations | Where-Object {
            ($_.Name -and $_.Name -ilike "*heska*") -or
            ($_.Description -and $_.Description -ilike "*heska*")
        }

        if ($heskaReservations) {
            foreach ($reservation in $heskaReservations) {
                $heskaDevices += [PSCustomObject]@{
                    Type        = "Reservation"
                    IPAddress   = $reservation.IPAddress
                    ClientId    = $reservation.ClientId
                    HostName    = $reservation.Name
                    Description = $reservation.Description
                    LeaseExpiry = "N/A (Reservation)"
                }
            }
        }
    } catch {
        Write-Host "Error fetching reservations from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Heska search"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Remove duplicates based on IPAddress (keep the first occurrence)
    $uniqueHeskaDevices = $heskaDevices | Sort-Object IPAddress -Unique

    # Display results
    if ($uniqueHeskaDevices.Count -eq 0) {
        Write-Host "No Heska devices found in DHCP leases or reservations for scope $scopeId on $selectedServer." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($uniqueHeskaDevices.Count) unique Heska device(s) in scope $scopeId on ${selectedServer} (duplicates removed)." -ForegroundColor Green
        $sortedDevices = $uniqueHeskaDevices | Sort-Object IPAddress
        $selectedDevice = $sortedDevices | Out-GridView -Title "#12 Search for Heska devices - v.$version - AU $AU - $(Get-Date -Format 'dddd, MMMM dd, yyyy  h:mm:ss tt')" -OutputMode Single

        # Ping options
        if ($uniqueHeskaDevices.Count -gt 0) {
            Write-Host "Ping Options:" -ForegroundColor Cyan
            Write-Host "1. Ping selected device" -ForegroundColor White
            Write-Host "2. Ping all found devices" -ForegroundColor White
            Write-Host "3. No ping" -ForegroundColor White
            $pingOption = (Read-Host "Choose an option (1-3)").Trim()

            switch ($pingOption) {
                "1" {
                    if ($selectedDevice) {
                        Write-Host "Pinging selected device $($selectedDevice.IPAddress)..." -ForegroundColor Cyan
                        try {
                            $pingResult = Test-Connection -ComputerName $selectedDevice.IPAddress -Count 4 -ErrorAction Stop
                            Write-Host "Ping Results for $($selectedDevice.IPAddress):" -ForegroundColor Green
                            $pingResult | Select-Object Address, ResponseTime, Status | Format-Table -AutoSize
                        } catch {
                            Write-Host "Ping failed for $($selectedDevice.IPAddress): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } else {
                        Write-Host "No device selected. Skipping ping." -ForegroundColor Yellow
                    }
                }
                "2" {
                    # Ping all found devices using the desired RSJob logic
                    $devicesToPing = $sortedDevices
                    $devicesToPingCount = $devicesToPing.Count
                    Write-Host "Pinging all $devicesToPingCount found devices in parallel..." -ForegroundColor Cyan

                    # The logic you provided for option 1 uses $otherDevices, but here we use $sortedDevices
                    $pingJobs = @()
                    foreach ($device in $devicesToPing) {
                        $job = Start-RSJob -ScriptBlock {
                            param($device)
                            $ip = $device.IPAddress.ToString()
                            # Use HostName first, then IPAddress for the name
                            $deviceName = if ($device.PSObject.Properties.Match('HostName') -and $device.HostName -and $device.HostName -ne $ip -and $device.HostName -notmatch '^BAD_ADDRESS$') {
                                $device.HostName
                            } else {
                                $ip
                            }
                            # Ping with Count 2 for efficiency, -Quiet returns True/False
                            $pingResult = Test-Connection -ComputerName $ip -Count 2 -Quiet
                            [PSCustomObject]@{
                                DeviceName = $deviceName
                                IP         = $ip
                                PingResult = $pingResult
                            }
                        } -ArgumentList $device
                        $pingJobs += $job
                    }
                    
                    # Wait for all ping jobs and display results
                    Write-Host "Waiting for ping results..." -ForegroundColor Cyan
                    $pingResults = $pingJobs | Wait-RSJob | Receive-RSJob
                    
                    foreach ($result in $pingResults | Sort-Object DeviceName) {
                        Write-Host "Device " -NoNewline
                        Write-Host "$($result.DeviceName)" -ForegroundColor Cyan -NoNewline
                        Write-Host " ($($result.IP)) : Ping - " -NoNewline
                        if ($result.PingResult) {
                            Write-Host "$($result.PingResult)" -ForegroundColor Green
                        } else {
                            Write-Host "$($result.PingResult)" -ForegroundColor Red
                        }
                    }
                    Remove-RSJob -Job $pingJobs
                }
                "3" {
                    Write-Host "No ping selected." -ForegroundColor Yellow
                }
                default {
                    Write-Host "Invalid option. No ping performed." -ForegroundColor Yellow
                }
            }
        }

        # Display selected device details if selected
        if ($selectedDevice) {
            Write-Host "Selected Device Details:" -ForegroundColor Cyan
            $selectedDevice | Format-List Type, IPAddress, ClientId, HostName, Description, LeaseExpiry
        }
    }

    # Export results if any found
    if ($uniqueHeskaDevices.Count -gt 0) {
        Export-Results -Results $uniqueHeskaDevices -BaseName "heska_devices" -AU $AU
    }

    Write-ConditionalLog "Heska Device Search completed for AU $AU (Server: $selectedServer, Scope: $scopeId)"
}

function ShowAllDHCPLeases {
    param([string]$AU)

    Write-Host "Showing all DHCP leases and reservations for AU $AU" -ForegroundColor Cyan
    Write-ConditionalLog "Starting Show All DHCP Leases for AU $AU"

    # Load admin credentials for DHCP server access (required for permissions)
    $adminCredPath = "$global:ScriptRoot\Private\vcaadmin.xml"
    if (Test-Path $adminCredPath) {
        try {
            $AdminCredential = Import-Clixml -Path $adminCredPath
            Write-ConditionalLog "Admin credentials loaded for DHCP access in AU $AU"
        } catch {
            Write-Host "Failed to load admin credentials from $adminCredPath : $($_.Exception.Message). DHCP access may fail." -ForegroundColor Red
            Write-ConditionalLog "Failed to load admin credentials for DHCP: $($_.Exception.Message)"
            $AdminCredential = $null
        }
    } else {
        Write-Host "Admin credentials file not found at $adminCredPath. Update via menu option 11. DHCP access may fail." -ForegroundColor Yellow
        Write-ConditionalLog "Admin credentials file missing for DHCP access in AU $AU"
        $AdminCredential = $null
    }

    # Get DHCP servers for this AU
    $dhcpServers = Get-DHCPServersForAU -AU $AU -Credential $ADCredential
    if (-not $dhcpServers -or $dhcpServers.Count -eq 0) {
        Write-Host "No DHCP servers configured or discovered for AU $AU. Skipping search." -ForegroundColor Yellow
        return
    }

    # Get scopes from all DHCP servers, filtered for this AU (like VCAHospLauncher option 24b)
    Write-Host "Retrieving DHCP scopes from all servers for AU $AU..." -ForegroundColor Cyan
    $allScopes = @()
    foreach ($server in $dhcpServers) {
        try {
            $serverScopes = Invoke-Command -ComputerName $server -ScriptBlock { param($au) Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$au-*" } } -ArgumentList $AU -Credential $AdminCredential -ErrorAction Stop
            if ($serverScopes) {
                # Add server info to each scope object
                foreach ($scope in $serverScopes) {
                    $scope | Add-Member -MemberType NoteProperty -Name "DHCPServer" -Value $server -Force
                    $allScopes += $scope
                }
            }
            Write-Host "Found $($serverScopes.Count) scopes on $server." -ForegroundColor Green
        } catch {
            Write-Host "Error fetching scopes from $server : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($allScopes.Count -eq 0) {
        Write-Host "No DHCP scopes found on any DHCP server for AU $AU." -ForegroundColor Red
        return
    }

    Write-Host "Found $($allScopes.Count) total scopes across all DHCP servers for AU $AU." -ForegroundColor Green

    # Select a scope from all available scopes (deduplicate by ScopeId)
    if ($allScopes.Count -eq 1) {
        $selectedScope = $allScopes[0]
        Write-Host "Using DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    } else {
        # Deduplicate scopes by ScopeId, keeping the first occurrence (first server found)
        $uniqueScopes = $allScopes | Group-Object -Property ScopeId | ForEach-Object {
            $_.Group | Select-Object -First 1
        }
        $scopeDisplay = $uniqueScopes | Select-Object Name, ScopeId, StartRange, EndRange, Description, LeaseDuration, SubnetMask, DHCPServer
        $selectedScope = $scopeDisplay | Out-GridView -Title "Select DHCP Scope for AU $AU" -OutputMode Single
        if (-not $selectedScope) {
            Write-Host "No DHCP scope selected. Operation cancelled." -ForegroundColor Yellow
            return
        }
        # Get the full scope object back from the original scopes array
        $selectedScope = $allScopes | Where-Object { $_.ScopeId -eq $selectedScope.ScopeId } | Select-Object -First 1
        Write-Host "Selected DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    }

    $scopeId = $selectedScope.ScopeId
    $selectedServer = $selectedScope.DHCPServer
    Write-Host "Selected scope: $scopeId ($($selectedScope.Name)) on server $selectedServer" -ForegroundColor Green

    $allDevices = @()

    # Search the selected DHCP server and scope
    Write-Host "Retrieving DHCP leases and reservations from $selectedServer for scope: $scopeId" -ForegroundColor Cyan

    # Get DHCP leases
    try {
        $leases = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Lease -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($leases.Count) DHCP leases from $selectedServer." -ForegroundColor Cyan

        foreach ($lease in $leases) {
            $allDevices += [PSCustomObject]@{
                Type        = "Lease"
                IPAddress   = $lease.IPAddress
                ClientId    = $lease.ClientId
                HostName    = $lease.HostName
                Description = $lease.Description
                LeaseExpiry = $lease.LeaseExpiryTime
            }
        }
    } catch {
        Write-Host "Error fetching leases from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Show All DHCP Leases"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Get DHCP reservations
    try {
        $reservations = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Reservation -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($reservations.Count) DHCP reservations from $selectedServer." -ForegroundColor Cyan

        foreach ($reservation in $reservations) {
            $allDevices += [PSCustomObject]@{
                Type        = "Reservation"
                IPAddress   = $reservation.IPAddress
                ClientId    = $reservation.ClientId
                HostName    = $reservation.Name
                Description = $reservation.Description
                LeaseExpiry = "N/A (Reservation)"
            }
        }
    } catch {
        Write-Host "Error fetching reservations from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Show All DHCP Leases"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Display results
    if ($allDevices.Count -eq 0) {
        Write-Host "No DHCP leases or reservations found in scope $scopeId on $selectedServer." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($allDevices.Count) DHCP leases and reservations in scope $scopeId on ${selectedServer}." -ForegroundColor Green
        $sortedDevices = $allDevices | Sort-Object IPAddress
        $selectedDevice = $sortedDevices | Out-GridView -Title "#12c Show All DHCP Leases - v.$version - AU $AU - $(Get-Date -Format 'dddd, MMMM dd, yyyy  h:mm:ss tt')" -OutputMode Single

        # Ping options
        if ($allDevices.Count -gt 0) {
            Write-Host "Ping Options:" -ForegroundColor Cyan
            Write-Host "1. Ping selected device" -ForegroundColor White
            Write-Host "2. Ping all found devices" -ForegroundColor White
            Write-Host "3. No ping" -ForegroundColor White
            $pingOption = (Read-Host "Choose an option (1-3)").Trim()

            switch ($pingOption) {
                "1" {
                    if ($selectedDevice) {
                        Write-Host "Pinging selected device $($selectedDevice.IPAddress)..." -ForegroundColor Cyan
                        try {
                            $pingResult = Test-Connection -ComputerName $selectedDevice.IPAddress -Count 4 -ErrorAction Stop
                            Write-Host "Ping Results for $($selectedDevice.IPAddress):" -ForegroundColor Green
                            $pingResult | Select-Object Address, ResponseTime, Status | Format-Table -AutoSize
                        } catch {
                            Write-Host "Ping failed for $($selectedDevice.IPAddress): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } else {
                        Write-Host "No device selected. Skipping ping." -ForegroundColor Yellow
                    }
                }
                "2" {
                    Write-Host "Pinging all $($allDevices.Count) found devices in parallel..." -ForegroundColor Cyan
                    
                    # Use runspaces for parallel ping testing (limit concurrency to prevent network saturation)
                    $maxConcurrency = 5  # Lowered from 10 to reduce "lack of resources" errors; adjust as needed
                    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrency)
                    $runspacePool.Open()
                    
                    $totalDevices = $sortedDevices.Count
                    $i = 0
                    $handles = @()
                    $powershellInstances = @()
                    
                    # Start parallel ping operations
                    foreach ($device in $sortedDevices) {
                        $i++
                        Write-Progress -Activity "Pinging devices" -Status "Device $i of $totalDevices : $($device.IPAddress)" -PercentComplete (($i / $totalDevices) * 100)
                        
                        $ps = [powershell]::Create()
                        $ps.RunspacePool = $runspacePool
                        
                        $scriptBlock = {
                            param($ipAddress)
                            try {
                                $pingResult = Test-Connection -ComputerName $ipAddress -Count 4 -ErrorAction Stop
                                $formattedOutput = $pingResult | Select-Object Address, ResponseTime, 
                                    @{Name='Status'; Expression={
                                        if ($_.StatusCode -eq 0) { 'Success' } 
                                        elseif ($_.StatusCode -eq 11010) { 'TimedOut' } 
                                        else { "Failed (Code: $($_.StatusCode))" }
                                    }} | Format-Table -AutoSize | Out-String
                                [PSCustomObject]@{
                                    IPAddress = $ipAddress
                                    Success   = $true
                                    Output    = $formattedOutput.Trim()
                                }
                            } catch {
                                [PSCustomObject]@{
                                    IPAddress = $ipAddress
                                    Success   = $false
                                    Error     = $_.Exception.Message
                                }
                            }
                        }
                        
                        # Suppress output from method calls
                        $null = $ps.AddScript($scriptBlock)
                        $null = $ps.AddArgument($device.IPAddress)
                        
                        $handle = $ps.BeginInvoke()
                        $handles += $handle
                        $powershellInstances += $ps
                    }
                    
                    # Wait for all pings to complete
                    while ($handles | Where-Object { -not $_.IsCompleted }) {
                        Start-Sleep -Milliseconds 100
                    }
                    
                    Write-Progress -Activity "Pinging devices" -Completed
                    
                    # Collect and display results
                    for ($j = 0; $j -lt $powershellInstances.Count; $j++) {
                        $ps = $powershellInstances[$j]
                        $handle = $handles[$j]
                        try {
                            $result = $ps.EndInvoke($handle)
                            if ($result.Success) {
                                Write-Host "Ping Results for $($result.IPAddress):" -ForegroundColor Green
                                Write-Host $result.Output
                            } else {
                                Write-Host "Ping failed for $($result.IPAddress): $($result.Error)" -ForegroundColor Red
                            }
                        } catch {
                            Write-Host "Error processing ping result for device $($sortedDevices[$j].IPAddress): $($_.Exception.Message)" -ForegroundColor Red
                        }
                        $ps.Dispose()
                    }
                    
                    # Close runspace pool
                    $runspacePool.Close()
                    $runspacePool.Dispose()
                }
                "3" {
                    Write-Host "No ping selected." -ForegroundColor Yellow
                }
                default {
                    Write-Host "Invalid option. No ping performed." -ForegroundColor Yellow
                }
            }
        }

        # Display selected device details if selected
        if ($selectedDevice) {
            Write-Host "Selected Device Details:" -ForegroundColor Cyan
            $selectedDevice | Format-List Type, IPAddress, ClientId, HostName, Description, LeaseExpiry
        }
    }

    # Export results if any found
    if ($allDevices.Count -gt 0) {
        Export-Results -Results $allDevices -BaseName "all_dhcp_leases" -AU $AU
    }

    Write-ConditionalLog "Show All DHCP Leases completed for AU $AU (Server: $selectedServer, Scope: $scopeId)"
}

function SearchCreditCardDevices {
    param([string]$AU)

    Write-Host "Searching for credit card devices for AU $AU" -ForegroundColor Cyan
    Write-ConditionalLog "Starting Credit Card Device Search for AU $AU"

    # Load admin credentials for DHCP server access (required for permissions)
    $adminCredPath = "$global:ScriptRoot\Private\vcaadmin.xml"
    if (Test-Path $adminCredPath) {
        try {
            $AdminCredential = Import-Clixml -Path $adminCredPath
            Write-ConditionalLog "Admin credentials loaded for DHCP access in AU $AU"
        } catch {
            Write-Host "Failed to load admin credentials from $adminCredPath : $($_.Exception.Message). DHCP access may fail." -ForegroundColor Red
            Write-ConditionalLog "Failed to load admin credentials for DHCP: $($_.Exception.Message)"
            $AdminCredential = $null
        }
    } else {
        Write-Host "Admin credentials file not found at $adminCredPath. Update via menu option 11. DHCP access may fail." -ForegroundColor Yellow
        Write-ConditionalLog "Admin credentials file missing for DHCP access in AU $AU"
        $AdminCredential = $null
    }

    # Get DHCP servers for this AU
    $dhcpServers = Get-DHCPServersForAU -AU $AU -Credential $ADCredential
    if (-not $dhcpServers -or $dhcpServers.Count -eq 0) {
        Write-Host "No DHCP servers configured or discovered for AU $AU. Skipping search." -ForegroundColor Yellow
        return
    }

    # Get scopes from all DHCP servers, filtered for this AU (like VCAHospLauncher option 24b)
    Write-Host "Retrieving DHCP scopes from all servers for AU $AU..." -ForegroundColor Cyan
    $allScopes = @()
    foreach ($server in $dhcpServers) {
        try {
            $serverScopes = Invoke-Command -ComputerName $server -ScriptBlock { param($au) Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$au-*" } } -ArgumentList $AU -Credential $AdminCredential -ErrorAction Stop
            if ($serverScopes) {
                # Add server info to each scope object
                foreach ($scope in $serverScopes) {
                    $scope | Add-Member -MemberType NoteProperty -Name "DHCPServer" -Value $server -Force
                    $allScopes += $scope
                }
            }
            Write-Host "Found $($serverScopes.Count) scopes on $server." -ForegroundColor Green
        } catch {
            Write-Host "Error fetching scopes from $server : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($allScopes.Count -eq 0) {
        Write-Host "No DHCP scopes found on any DHCP server for AU $AU." -ForegroundColor Red
        return
    }

    Write-Host "Found $($allScopes.Count) total scopes across all DHCP servers for AU $AU." -ForegroundColor Green

    # Select a scope from all available scopes (deduplicate by ScopeId)
    if ($allScopes.Count -eq 1) {
        $selectedScope = $allScopes[0]
        Write-Host "Using DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    } else {
        # Deduplicate scopes by ScopeId, keeping the first occurrence (first server found)
        $uniqueScopes = $allScopes | Group-Object -Property ScopeId | ForEach-Object {
            $_.Group | Select-Object -First 1
        }
        $scopeDisplay = $uniqueScopes | Select-Object Name, ScopeId, StartRange, EndRange, Description, LeaseDuration, SubnetMask, DHCPServer
        $selectedScope = $scopeDisplay | Out-GridView -Title "Select DHCP Scope for AU $AU" -OutputMode Single
        if (-not $selectedScope) {
            Write-Host "No DHCP scope selected. Operation cancelled." -ForegroundColor Yellow
            return
        }
        # Get the full scope object back from the original scopes array
        $selectedScope = $allScopes | Where-Object { $_.ScopeId -eq $selectedScope.ScopeId } | Select-Object -First 1
        Write-Host "Selected DHCP scope: $($selectedScope.ScopeId) ($($selectedScope.Name)) on $($selectedScope.DHCPServer)" -ForegroundColor Cyan
    }

    $scopeId = $selectedScope.ScopeId
    $selectedServer = $selectedScope.DHCPServer
    Write-Host "Selected scope: $scopeId ($($selectedScope.Name)) on server $selectedServer" -ForegroundColor Green

    $creditCardDevices = @()

    # Search the selected DHCP server and scope
    Write-Host "Searching DHCP server: $selectedServer for scope: $scopeId" -ForegroundColor Cyan

    # Get DHCP leases
    try {
        $leases = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Lease -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($leases.Count) DHCP leases from $selectedServer." -ForegroundColor Cyan

        # Filter for credit card devices in leases
        $creditCardLeases = $leases | Where-Object {
            ($_.HostName -and $_.HostName -ilike "*.cc.vca.com") -or
            ($_.Description -and $_.Description -ilike "*.cc.vca.com")
        }

        if ($creditCardLeases) {
            foreach ($lease in $creditCardLeases) {
                $creditCardDevices += [PSCustomObject]@{
                    Type        = "Lease"
                    IPAddress   = $lease.IPAddress
                    ClientId    = $lease.ClientId
                    HostName    = $lease.HostName
                    Description = $lease.Description
                    LeaseExpiry = $lease.LeaseExpiryTime
                }
            }
        }
    } catch {
        Write-Host "Error fetching leases from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Credit Card search"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Get DHCP reservations
    try {
        $reservations = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($s) Get-DhcpServerv4Reservation -ScopeId $s } -ArgumentList $scopeId -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Fetched $($reservations.Count) DHCP reservations from $selectedServer." -ForegroundColor Cyan

        # Filter for credit card devices in reservations
        $creditCardReservations = $reservations | Where-Object {
            ($_.Name -and $_.Name -ilike "*.cc.vca.com") -or
            ($_.Description -and $_.Description -ilike "*.cc.vca.com")
        }

        if ($creditCardReservations) {
            foreach ($reservation in $creditCardReservations) {
                $creditCardDevices += [PSCustomObject]@{
                    Type        = "Reservation"
                    IPAddress   = $reservation.IPAddress
                    ClientId    = $reservation.ClientId
                    HostName    = $reservation.Name
                    Description = $reservation.Description
                    LeaseExpiry = "N/A (Reservation)"
                }
            }
        }
    } catch {
        Write-Host "Error fetching reservations from $selectedServer : $($_.Exception.Message)" -ForegroundColor Yellow
        if ($_.Exception.Message -match "access denied|credential|authentication|logon failure|unauthorized|permission") {
            Write-Host "This appears to be a credential issue. Would you like to update the admin credentials? (y/n)" -ForegroundColor Yellow
            $updateCred = Read-Host
            if ($updateCred.ToLower() -eq 'y') {
                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\adminuser)"
                if ($newCred) {
                    $newCred | Export-Clixml -Path $adminCredPath -Force
                    Write-Host "Admin credentials updated." -ForegroundColor Green
                    Write-ConditionalLog "Admin credentials updated due to DHCP error in Credit Card search"
                    $AdminCredential = $newCred
                }
            }
        }
    }

    # Remove duplicates based on IPAddress (keep the first occurrence)
    $uniqueCreditCardDevices = $creditCardDevices | Sort-Object IPAddress -Unique

    # Display results
    if ($uniqueCreditCardDevices.Count -eq 0) {
        Write-Host "No credit card devices found in DHCP leases or reservations for scope $scopeId on $selectedServer." -ForegroundColor Yellow
    } else {
        Write-Host "Found $($uniqueCreditCardDevices.Count) unique credit card device(s) in scope $scopeId on ${selectedServer} (duplicates removed)." -ForegroundColor Green
        $sortedDevices = $uniqueCreditCardDevices | Sort-Object IPAddress
        $selectedDevice = $sortedDevices | Out-GridView -Title "#12d Search for credit card devices - v.$version - AU $AU - $(Get-Date -Format 'dddd, MMMM dd, yyyy  h:mm:ss tt')" -OutputMode Single

        # Ping options
        if ($uniqueCreditCardDevices.Count -gt 0) {
            Write-Host "Ping Options:" -ForegroundColor Cyan
            Write-Host "1. Ping selected device" -ForegroundColor White
            Write-Host "2. Ping all found devices" -ForegroundColor White
            Write-Host "3. No ping" -ForegroundColor White
            $pingOption = (Read-Host "Choose an option (1-3)").Trim()

            switch ($pingOption) {
                "1" {
                    if ($selectedDevice) {
                        Write-Host "Pinging selected device $($selectedDevice.IPAddress)..." -ForegroundColor Cyan
                        try {
                            $pingResult = Test-Connection -ComputerName $selectedDevice.IPAddress -Count 4 -ErrorAction Stop
                            Write-Host "Ping Results for $($selectedDevice.IPAddress):" -ForegroundColor Green
                            $pingResult | Select-Object Address, ResponseTime, Status | Format-Table -AutoSize
                        } catch {
                            Write-Host "Ping failed for $($selectedDevice.IPAddress): $($_.Exception.Message)" -ForegroundColor Red
                        }
                    } else {
                        Write-Host "No device selected. Skipping ping." -ForegroundColor Yellow
                    }
                }
                "2" {
                    # Ping all found devices using the desired RSJob logic
                    $devicesToPing = $sortedDevices
                    $devicesToPingCount = $devicesToPing.Count
                    Write-Host "Pinging all $devicesToPingCount found devices in parallel..." -ForegroundColor Cyan

                    # The logic you provided for option 1 uses $otherDevices, but here we use $sortedDevices
                    $pingJobs = @()
                    foreach ($device in $devicesToPing) {
                        $job = Start-RSJob -ScriptBlock {
                            param($device)
                            $ip = $device.IPAddress.ToString()
                            # Use HostName first, then IPAddress for the name
                            $deviceName = if ($device.PSObject.Properties.Match('HostName') -and $device.HostName -and $device.HostName -ne $ip -and $device.HostName -notmatch '^BAD_ADDRESS$') {
                                $device.HostName
                            } else {
                                $ip
                            }
                            # Ping with Count 2 for efficiency, -Quiet returns True/False
                            $pingResult = Test-Connection -ComputerName $ip -Count 2 -Quiet
                            [PSCustomObject]@{
                                DeviceName = $deviceName
                                IP         = $ip
                                PingResult = $pingResult
                            }
                        } -ArgumentList $device
                        $pingJobs += $job
                    }
                    
                    # Wait for all ping jobs and display results
                    Write-Host "Waiting for ping results..." -ForegroundColor Cyan
                    $pingResults = $pingJobs | Wait-RSJob | Receive-RSJob
                    
                    foreach ($result in $pingResults | Sort-Object DeviceName) {
                        Write-Host "Device " -NoNewline
                        Write-Host "$($result.DeviceName)" -ForegroundColor Cyan -NoNewline
                        Write-Host " ($($result.IP)) : Ping - " -NoNewline
                        if ($result.PingResult) {
                            Write-Host "$($result.PingResult)" -ForegroundColor Green
                        } else {
                            Write-Host "$($result.PingResult)" -ForegroundColor Red
                        }
                    }
                    Remove-RSJob -Job $pingJobs
                }
                "3" {
                    Write-Host "No ping selected." -ForegroundColor Yellow
                }
                default {
                    Write-Host "Invalid option. No ping performed." -ForegroundColor Yellow
                }
            }
        }

        # Display selected device details if selected
        if ($selectedDevice) {
            Write-Host "Selected Device Details:" -ForegroundColor Cyan
            $selectedDevice | Format-List Type, IPAddress, ClientId, HostName, Description, LeaseExpiry
        }
    }

    # Export results if any found
    if ($uniqueCreditCardDevices.Count -gt 0) {
        Export-Results -Results $uniqueCreditCardDevices -BaseName "credit_card_devices" -AU $AU
    }

    Write-ConditionalLog "Credit Card Device Search completed for AU $AU (Server: $selectedServer, Scope: $scopeId)"
}