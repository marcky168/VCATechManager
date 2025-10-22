# Network Scanner - Grok.ps1
# Version: 1.0
# Description: Scans the local network subnet for active devices by pinging IPs and retrieving MAC addresses from ARP cache.
# This can be run locally or remotely via Invoke-Command. Assumes /24 subnet. Run on the remote computer to scan its network.

# Function to scan the network and get IPs and MACs
function Scan-Network {
    [CmdletBinding()]
    param(
        [string]$SubnetMask = "/24"  # Assuming /24; can be extended for others
    )

    # Determine local subnet if not provided
    try {
        $ipConfig = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -First 1
        if (-not $ipConfig) {
            throw "No default IPv4 configuration found."
        }
        $localIP = $ipConfig.IPv4Address.IPAddress
        $subnet = $localIP -replace '\.\d+$', '.0/24'
        Write-Host "Scanning subnet: $subnet" -ForegroundColor Green
    } catch {
        Write-Error "Could not determine local subnet: $_"
        return
    }

    # Generate IP range (1-254, skipping .0 and .255)
    $baseSubnet = $subnet -replace '/24$', ''
    $ipRange = 1..254 | ForEach-Object { "$baseSubnet$_" }

    # Ping all IPs to find active hosts and populate ARP cache
    Write-Host "Pinging IPs... This may take a few minutes." -ForegroundColor Yellow
    $activeIPs = @()
    foreach ($ip in $ipRange) {
        if (Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            $activeIPs += $ip
            Write-Host "Active: $ip" -ForegroundColor Green
        }
    }

    if ($activeIPs.Count -eq 0) {
        Write-Host "No active devices found on the subnet." -ForegroundColor Red
        return
    }

    # Get ARP entries (Net Neighbors) for active IPs
    $arpEntries = Get-NetNeighbor -AddressFamily IPv4 | Where-Object { $_.State -eq 'Reachable' -or $_.State -eq 'Stale' }

    # Match active IPs to MACs
    $results = foreach ($ip in $activeIPs) {
        $entry = $arpEntries | Where-Object { $_.IPAddress -eq $ip }
        [PSCustomObject]@{
            IPAddress = $ip
            MACAddress = if ($entry) { $entry.LinkLayerAddress } else { "Unknown (ARP not resolved)" }
            Status = if ($entry) { $entry.State } else { "Pinged but no ARP" }
        }
    }

    # Display results
    $results | Sort-Object IPAddress | Format-Table -AutoSize
    Write-Host "`nScan complete. Found $($results.Count) active devices." -ForegroundColor Green

    return $results
}

