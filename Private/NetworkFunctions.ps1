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

    # First, select a DHCP server
    if ($dhcpServers.Count -eq 1) {
        $selectedServer = $dhcpServers[0]
        Write-Host "Using DHCP server: $selectedServer" -ForegroundColor Cyan
    } else {
        Write-Host "Available DHCP servers:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $dhcpServers.Count; $i++) {
            Write-Host "$($i + 1). $($dhcpServers[$i])" -ForegroundColor White
        }

        $serverChoice = Read-Host "Select DHCP server (1-$($dhcpServers.Count))"
        try {
            $serverIndex = [int]$serverChoice - 1
            if ($serverIndex -lt 0 -or $serverIndex -ge $dhcpServers.Count) {
                Write-Host "Invalid server selection. Operation cancelled." -ForegroundColor Yellow
                return
            }
            $selectedServer = $dhcpServers[$serverIndex]
        } catch {
            Write-Host "Invalid input. Operation cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Get scopes from the selected server, filtered for this AU
    Write-Host "Retrieving DHCP scopes from $selectedServer..." -ForegroundColor Cyan
    try {
        $scopes = Invoke-Command -ComputerName $selectedServer -ScriptBlock { param($au) Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$au-*" } } -ArgumentList $AU -Credential $AdminCredential -ErrorAction Stop
        Write-Host "Found $($scopes.Count) scopes for AU $AU on $selectedServer." -ForegroundColor Green
    } catch {
        Write-Host "Error fetching scopes from $selectedServer : $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This may be due to network connectivity, permissions, or server unavailability." -ForegroundColor Yellow
        return
    }

    if ($scopes.Count -eq 0) {
        Write-Host "No DHCP scopes found on $selectedServer." -ForegroundColor Red
        return
    }

    # Select a scope
    Write-Host "Available DHCP scopes for AU $AU on $selectedServer :" -ForegroundColor Cyan
    for ($i = 0; $i -lt $scopes.Count; $i++) {
        $scope = $scopes[$i]
        Write-Host "$($i + 1). $($scope.ScopeId) - $($scope.Name) ($($scope.StartRange) - $($scope.EndRange))" -ForegroundColor White
    }

    $scopeChoice = Read-Host "Select DHCP scope (1-$($scopes.Count))"
    try {
        $scopeIndex = [int]$scopeChoice - 1
        if ($scopeIndex -lt 0 -or $scopeIndex -ge $scopes.Count) {
            Write-Host "Invalid scope selection. Operation cancelled." -ForegroundColor Yellow
            return
        }
        $selectedScope = $scopes[$scopeIndex]
    } catch {
        Write-Host "Invalid input. Operation cancelled." -ForegroundColor Yellow
        return
    }

    $scopeId = $selectedScope.ScopeId
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
    }

    # Remove duplicates based on IPAddress (keep the first occurrence)
    $uniqueHeskaDevices = $heskaDevices | Sort-Object IPAddress -Unique

    # Display results
    if ($uniqueHeskaDevices.Count -eq 0) {
        Write-Host "No Heska devices found in DHCP leases or reservations for scope $scopeId on $selectedServer." -ForegroundColor Yellow
    } else {
        Write-Host "`nFound $($uniqueHeskaDevices.Count) unique Heska device(s) in scope $scopeId on $selectedServer (duplicates removed):" -ForegroundColor Green
        $uniqueHeskaDevices | Sort-Object IPAddress | Format-Table -AutoSize -Property Type, IPAddress, ClientId, HostName, Description, LeaseExpiry
    }

    # Export results if any found
    if ($uniqueHeskaDevices.Count -gt 0) {
        Export-Results -Results $uniqueHeskaDevices -BaseName "heska_devices" -AU $AU
    }

    Write-ConditionalLog "Heska Device Search completed for AU $AU (Server: $selectedServer, Scope: $scopeId)"
}