# Function for Device Connectivity Test (Option 12) - Independent version

function DeviceConnectivityTest {
    param([string]$AU)

    Write-Host "DeviceConnectivityTest function called for AU $AU" -ForegroundColor Cyan  # Debug: Confirm function is running
    Write-ConditionalLog "Starting Device Connectivity Test for AU $AU"

    # Independent DHCP lease fetching
    $dhcpServer = $global:config.NetworkSettings.PrimaryDHCPServer
    $hostname = Convert-VcaAu -AU $AU -Suffix '-gw'
    try {
        $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
        if (-not $ipAddresses -or $ipAddresses.Length -eq 0) {
            Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
            return
        }
        $ip = $ipAddresses[0].ToString()
        $scopeId = $ip -replace '\.\d+$', '.0'
        Write-Host "Resolved hostname '$hostname' to IP '$ip', using scope '$scopeId'." -ForegroundColor Cyan
        $leases = Get-DhcpServerv4Lease -ComputerName $dhcpServer -ScopeId $scopeId -ErrorAction Stop
        Write-Host "Fetched $($leases.Count) DHCP leases for scope $scopeId." -ForegroundColor Cyan
    } catch {
        Write-Host "Error in DNS or DHCP fetching: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    # Filter for Fuse devices (fixed: proper matching logic with error handling)
    $macPrefixes = @{
        "Fuse"  = @("00-90-FB", "00-50-56", "00-0C-29")
    }
    try {
        $fuseDevices = $leases | Where-Object {
            if (-not $_.ClientId) { return $false }
            $normalizedClientId = ($_.ClientId -replace "-", "").ToUpper()
            $match = $false
            foreach ($prefix in $macPrefixes["Fuse"]) {
                $normalizedPrefix = $prefix.Replace("-", "").ToUpper()
                if ($normalizedClientId.StartsWith($normalizedPrefix)) {
                    $match = $true
                    break
                }
            }
            $match
        }
        Write-Host "Found $($fuseDevices.Count) Fuse devices after filtering." -ForegroundColor Cyan
    } catch {
        Write-Host "Error during MAC filtering: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if (-not $fuseDevices -or $fuseDevices.Count -eq 0) {
        Write-Host "No Fuse devices found in DHCP leases. Check MAC prefixes or scope." -ForegroundColor Yellow
        return
    }

    foreach ($device in $fuseDevices) {
        try {
            $ip = $device.IPAddress.ToString()  # Ensure string for compatibility
            $ping = Test-Connection -ComputerName $ip -Count 2 -Quiet
            # Simplified port check (no job for reliability)
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $connectResult = $tcpClient.BeginConnect($ip, 8443, $null, $null)
            $connected = $connectResult.AsyncWaitHandle.WaitOne(5000)
            $tcpClient.Close()
            $port8443 = $connected
            Write-Host "Device $ip : Ping - $ping, Port 8443 - $port8443" -ForegroundColor Green
        } catch {
            Write-Host "Error testing connectivity for device $($device.IPAddress): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-ConditionalLog "Device Connectivity Test completed for AU $AU"
}
