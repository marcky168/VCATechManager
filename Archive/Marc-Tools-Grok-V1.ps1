# Combined PowerShell Script with Menu Options

# Set version
$version = "1"

# Get script path and last write time
$scriptPath = $MyInvocation.MyCommand.Path
if ($scriptPath) {
    $lastWritten = (Get-Item $scriptPath).LastWriteTime.ToString("MM/dd/yyyy HH:mm:ss")
} else {
    $lastWritten = "N/A"
}

# Import ActiveDirectory module
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Import required modules
Import-Module -Name "$PSScriptRoot\Private\lib\PoshRSJob" -ErrorAction SilentlyContinue
Import-Module -Name "$PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue

# Dot-source functions from Private folder
. "$PSScriptRoot\Private\Convert-VcaAU.ps1"
. "$PSScriptRoot\Private\Get-DiskUsage.ps1"
. "$PSScriptRoot\Private\Get-DriveStatus.ps1"
. "$PSScriptRoot\Private\Get-FirmwareVersion.ps1"
. "$PSScriptRoot\Private\Get-MemoryUsage.ps1"
. "$PSScriptRoot\Private\Get-OldUserProfiles.ps1"
. "$PSScriptRoot\Private\Get-OldVhds.ps1"
. "$PSScriptRoot\Private\Get-RdsConnectionConfig.ps1"
. "$PSScriptRoot\Private\Get-UPSStatus.ps1"
. "$PSScriptRoot\Private\Get-VCAHeadCount.ps1"
. "$PSScriptRoot\Private\Get-VCAHPEDriveFirmwareInfo.ps1"
. "$PSScriptRoot\Private\New-ServiceNowGUI.ps1"
. "$PSScriptRoot\Private\New-ServiceNowIncident.ps1"
. "$PSScriptRoot\Private\Remove-BakRegistry.ps1"
. "$PSScriptRoot\Private\Version.ps1"
. "$PSScriptRoot\Private\whatdisk.ps1"
. "$PSScriptRoot\Private\whatusers.ps1"

# Set window title
$host.UI.RawUI.WindowTitle = "Marc Tools V1 - Grok - $scriptPath (Written:$lastWritten)"

# Set console colors to match the style (dark blue background, white foreground)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# Display portal title in cyan
Write-Host "Marc Tools V1 - Grok" -ForegroundColor Cyan

# Function to get servers for AU (modified to only include -NS servers)
function Get-VCAServers {
    param(
        [string]$AU
    )

    if ($AU -notmatch '^\d{3,6}$') {
        throw "Invalid AU number. Please enter a 3 to 6 digit number."
    }

    $SiteAU = Convert-VcaAu -AU $AU -Suffix ''

    # Filter specifically for -ns* servers
    $adServers = Get-ADComputer -Filter "Name -like '$SiteAU-ns*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties Name | Select-Object -ExpandProperty Name | Sort-Object Name

    if (-not $adServers) {
        throw "No -NS servers found for AU $AU."
    }

    return $adServers
}

# Function for Abaxis MAC Address Search
function Abaxis-MacAddressSearch {
    param([string]$AU)

    # Function to normalize MAC addresses by removing hyphens and converting to uppercase
    function Normalize-MacAddressForDhcpScript {
        param ($mac)
        return $mac.Replace("-", "").ToUpper()
    }

    # Function to determine the device group based on MAC address
    function Get-DeviceGroup {
        param ($mac)
        $normalizedMac = Normalize-MacAddressForDhcpScript $mac
        foreach ($group in $macPrefixes.Keys) {
            $prefixes = $macPrefixes[$group]
            $normalizedPrefixes = $prefixes | ForEach-Object { Normalize-MacAddressForDhcpScript $_ }
            if ($normalizedPrefixes | Where-Object { $normalizedMac.StartsWith($_) }) {
                return $group
            }
        }
        return "Other"
    }

    # Define MAC address prefixes for each device group
    $macPrefixes = @{
        "VS2"   = @("00-07-32", "00-30-64")
        "HM5"   = @("00-1B-EB")
        "VSPro" = @("00-03-1D")
        "Fuse"  = @("00-90-FB", "00-50-56", "00-0C-29")
    }

    # DHCP server hostname
    $dhcpServer = "phhospdhcp2.vcaantech.com"

    $hostname = Convert-VcaAU -AU $AU -Suffix '-gw'

    # Resolve hostname to IP address with error handling
    try {
        $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
    } catch {
        Write-Host "Error: Could not resolve hostname '$hostname'. $_" -ForegroundColor Red
        return
    }

    if ($ipAddresses.Length -eq 0) {
        Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
        return
    } elseif ($ipAddresses.Length -gt 1) {
        Write-Host "Warning: Multiple IP addresses found for '$hostname'. Using the first one: $($ipAddresses[0].ToString())" -ForegroundColor Yellow
    }
    $ip = $ipAddresses[0].ToString()

    # Calculate scope ID assuming a /24 subnet (e.g., 192.168.1.0)
    $scopeId = $ip -replace '\.\d+$', '.0'

    # Retrieve DHCP leases for the scope with error handling
    try {
        $leases = Get-DhcpServerv4Lease -ComputerName $dhcpServer -ScopeId $scopeId -ErrorAction Stop
    } catch {
        Write-Host "Error: Could not retrieve leases from DHCP server '$dhcpServer'. $_" -ForegroundColor Red
        return
    }

    if (-not $leases) {
        Write-Host "No leases found for scope '$scopeId'."
    }

    # Process each group and find matching leases
    foreach ($group in $macPrefixes.Keys) {
        $prefixes = $macPrefixes[$group]
        $normalizedPrefixes = $prefixes | ForEach-Object { Normalize-MacAddressForDhcpScript $_ }

        $matchingLeases = $leases | Where-Object {
            $normalizedClientId = Normalize-MacAddressForDhcpScript $_.ClientId
            $normalizedPrefixes | Where-Object { $normalizedClientId.StartsWith($_) }
        }

        if ($matchingLeases) {
            Write-Host "`nLeases for $group" -ForegroundColor Green
            $matchingLeases | Sort-Object IPAddress | Format-Table -Property IPAddress, ClientId, @{Name="LastLeased"; Expression={$_.LeaseExpiryTime}}
        } else {
            Write-Host "`nNo leases found for group $group"
        }
    }

    # Retrieve and display DHCP reservations for the scope
    Write-Host "`nDHCP Reservations for scope $scopeId" -ForegroundColor Green
    try {
        $reservations = Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scopeId -ErrorAction Stop
        if ($reservations) {
            $reservations | Sort-Object IPAddress | Format-Table -Property IPAddress, ClientId, Name, Description, @{Name="DeviceGroup"; Expression={Get-DeviceGroup $_.ClientId}}
        } else {
            Write-Host "No reservations found for scope '$scopeId'."
        }
    } catch {
        Write-Host "Error: Could not retrieve reservations from DHCP server '$dhcpServer'. $_" -ForegroundColor Red
    }

    # Add nslookup for Hxxxx-fuse
    $fuseHostname = Convert-VcaAU -AU $AU -Suffix '-fuse'
    try {
        $fuseIpAddresses = [System.Net.Dns]::GetHostAddresses($fuseHostname)
        if ($fuseIpAddresses.Length -gt 0) {
            $fuseIp = $fuseIpAddresses[0].ToString()
            Write-Host "`nFuse Device IP (from nslookup on $fuseHostname): $fuseIp" -ForegroundColor Green

            # Ping the Fuse IP
            Write-Host "`nPinging Fuse device at $fuseIp..." -ForegroundColor Cyan
            $pingResult = Test-Connection -ComputerName $fuseIp -Count 4 -ErrorAction SilentlyContinue
            if ($pingResult) {
                $pingResult | Format-Table -Property Address, ResponseTime, StatusCode
                Write-Host "Fuse device is responsive." -ForegroundColor Green

                # If ping successful, open the Fuse webpage
                $fuseUrl = "https://$fuseHostname`:8443"
                Start-Process $fuseUrl
                Write-Host "Opening Fuse webpage: $fuseUrl" -ForegroundColor Green
            } else {
                Write-Host "Fuse device did not respond to ping." -ForegroundColor Red
            }
        } else {
            Write-Host "`nNo IP found for Fuse device ($fuseHostname)." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error resolving Fuse hostname '$fuseHostname': $_" -ForegroundColor Red
    }
}

# Function for Woofware Errors Check
function Woofware-ErrorsCheck {
    param([string]$AU)

    try {
        $servers = Get-VCAServers -AU $AU
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return
    }

    ForEach ($server in $servers) {
        Write-Host "Querying server: $server" -ForegroundColor Yellow
        Invoke-Command -ComputerName $server -ScriptBlock {
            $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
            $serverTime = $using:server + '  ' + $time
            Write-Host $serverTime -ForegroundColor Cyan

            $events101103 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=101,103} –MaxEvents 50 -ErrorAction SilentlyContinue
            if ($events101103) {
                Write-Host "Events 101 and 103:" -ForegroundColor Green
                $events101103 | Format-Table TimeCreated, Id, LevelDisplayName, Message
            } else {
                Write-Host "No events 101 or 103 found." -ForegroundColor Yellow
            }

            $event103 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=103} –MaxEvents 1 -ErrorAction Ignore
            if ($event103) {
                Write-Host "Latest Event 103:" -ForegroundColor Green
                $event103 | Format-List TimeCreated, Id, LevelDisplayName, Message
            } else {
                Write-Host "No event 103 found." -ForegroundColor Yellow
            }

            $events101 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=101} –MaxEvents 2 -ErrorAction Ignore
            if ($events101) {
                Write-Host "Latest 2 Events 101:" -ForegroundColor Green
                $events101 | Format-List TimeCreated, Id, LevelDisplayName, Message
            } else {
                Write-Host "No events 101 found." -ForegroundColor Yellow
            }

            $events100 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=100} –MaxEvents 10 -ErrorAction Ignore
            if ($events100) {
                Write-Host "Latest 10 Events 100:" -ForegroundColor Green
                $events100 | Format-List TimeCreated, Id, LevelDisplayName, Message
            } else {
                Write-Host "No events 100 found." -ForegroundColor Yellow
            }
        }
    }
}

# Function for Add DHCP Reservation
function Add-DHCPReservation {
    param([string]$AU)

    # Set variables for the DHCP servers and reservation name
    $DHCPServers = "phhospdhcp1.vcaantech.com", "phhospdhcp2.vcaantech.com"
    $ReservationName = "fuse"

    # Format the hostname based on AU
    $HospitalNumber = $AU
    if ($HospitalNumber.Length -eq 3) {
        $HospitalNumber = "0" + $HospitalNumber
    }
    $hostname = "h" + $HospitalNumber + "-gw"

    # Prompt for the MAC address suffix
    Write-Host "Enter the MAC address characters after 00-90-FB: " -NoNewline -ForegroundColor Cyan
    $MACAddressSuffix = Read-Host
    $MACAddress = "00-90-FB-" + $MACAddressSuffix

    # Resolve hostname to IP address
    $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)

    if ($ipAddresses.Length -eq 0) {
        Write-Host "Error: Could not resolve hostname '$hostname'." -ForegroundColor Red
        return
    }

    $ip = $ipAddresses[0].IPAddressToString
    $scopeId = $ip -replace '\.[0-9]+$', '.0'
    $scopeId = [System.Net.IPAddress]::Parse($scopeId)
    $ReservationIP = $ip -replace '.[0-9]+$', '.210'

    # Add the DHCP reservation to the specified scope on each server
    foreach ($Server in $DHCPServers) {
        # Check if the DHCP reservation already exists
        $ExistingReservation = Get-DhcpServerv4Reservation -ComputerName $Server -IPaddress $ReservationIP 
        if ($ExistingReservation) {
            # Prompt the user to confirm deletion
            Write-Host "A DHCP reservation with IP address $ReservationIP and scope $ScopeId already exists on server $Server. Do you want to delete it? (y/n): " -NoNewline -ForegroundColor Yellow
            $Confirm = Read-Host
            if ($Confirm -eq "y") {
                # Delete the existing reservation
                try {
                  Remove-DhcpServerv4Reservation -ComputerName $Server -IPAddress $ReservationIP -ErrorAction Stop
                  Write-Output "Deleted DHCP reservation for IP address $ReservationIP and scope $ScopeId on server $Server" -ForegroundColor Green
                  # Add the DHCP reservation
                  Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for $ReservationName"
                  Write-Output "Added DHCP reservation for IP address $ReservationIP to scope $ScopeId on server $Server" -ForegroundColor Green
                }
                catch {
                    Write-Output "Error deleting DHCP reservation for IP address $ReservationIP and scope $ScopeId on server '$Server': $_" -ForegroundColor Red
                }
            }
        }
        else {
            # Add the DHCP reservation
            Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for $ReservationName"
            Write-Output "Added DHCP reservation for IP address $ReservationIP to scope $ScopeId on server $Server" -ForegroundColor Green
        }
    }
}

# Function for User Logon Check
function User-LogonCheck {
    param(
        [string]$Username,
        [string]$AU
    )

    # Prompt for username if not provided
    if (-not $Username) {
        Write-Host "Enter the username: " -NoNewline -ForegroundColor Cyan
        $Username = Read-Host
    }

    try {
        $servers = Get-VCAServers -AU $AU
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return
    }

    # Query each server and collect results
    $results = foreach ($server in $servers) {
        Write-Host "Querying server: $server" -ForegroundColor Yellow
        try {
            $result = Invoke-Command -ComputerName $server -ScriptBlock {
                param($user)
                # Define the XPath filter for Event ID 4624 and the target username
                $filterXPath = "*[System[EventID=4624] and EventData/Data[@Name='TargetUserName']='$user']"
                
                # Retrieve the last 5 logon events and add IP address as a property
                $events = Get-WinEvent -LogName Security -FilterXPath $filterXPath -MaxEvents 5 -ErrorAction SilentlyContinue | ForEach-Object {
                    $eventXml = [xml]$_.ToXml()
                    $ipAddress = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
                    $_ | Add-Member -NotePropertyName 'IpAddress' -NotePropertyValue $ipAddress -PassThru
                }

                if ($events) {
                    # Select the most recent event with a valid IP address
                    $selectedEvent = $events | Where-Object { $_.IpAddress -and $_.IpAddress -ne "-" } | Select-Object -First 1
                    if ($selectedEvent) {
                        [PSCustomObject]@{
                            Server = $env:COMPUTERNAME
                            TimeCreated = $selectedEvent.TimeCreated
                            IpAddress = $selectedEvent.IpAddress
                        }
                    } else {
                        # If no event has an IP address, use the most recent event with "N/A"
                        $firstEvent = $events | Select-Object -First 1
                        [PSCustomObject]@{
                            Server = $env:COMPUTERNAME
                            TimeCreated = $firstEvent.TimeCreated
                            IpAddress = "N/A"
                        }
                    }
                } else {
                    # No events found for the user
                    [PSCustomObject]@{
                        Server = $env:COMPUTERNAME
                        TimeCreated = "No logon events found for user '$user'"
                        IpAddress = "N/A"
                    }
                }
            } -ArgumentList $Username
            $result
        } catch {
            # Handle errors such as unreachable servers
            [PSCustomObject]@{
                Server = $server
                TimeCreated = "Error: $_"
                IpAddress = "N/A"
            }
        }
    }

    # Display the results in a table
    $results | Format-Table -AutoSize
}

# Function for List AD Users and Check Logon (Option 5)
function ListADUsersAndCheckLogon {
    param([string]$AU)

    $groupName = "H" + (Convert-VcaAU -AU $AU -Prefix '' -Suffix '')

    try {
        $group = Get-ADGroup -Identity $groupName
        $users = Get-ADGroupMember -Identity $group -Recursive | Get-ADUser -Properties Name, SamAccountName, Department, Title, City, State, telephoneNumber | Select-Object Name, SamAccountName, Department, Title, @{n='Location'; e={$_.City + ', ' + $_.State}}, telephoneNumber
    } catch {
        Write-Host "Error fetching AD group or members: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    if ($users) {
        $selected = $users | Out-GridView -Title "AD Users for AU $AU ($groupName)" -OutputMode Single

        if ($selected) {
            $selectedUsername = $selected.SamAccountName
            User-LogonCheck -Username $selectedUsername -AU $AU
        }
    } else {
        Write-Host "No users found in group $groupName." -ForegroundColor Yellow
    }
}

# New Function for Kill Sparky Shell (Option 6)
function Kill-SparkyShell {
    param([string]$AU)

    try {
        $servers = Get-VCAServers -AU $AU
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return
    }

    # Collect active sessions from all servers
    $sessions = foreach ($server in $servers) {
        Write-Host "Querying active sessions on server: $server" -ForegroundColor Yellow
        try {
            Get-TSSession -ComputerName $server -State Active -ErrorAction SilentlyContinue | 
            Where-Object { $_.UserName } |  # Filter out sessions without usernames
            Select-Object @{n='Server'; e={$server}}, UserName, SessionId, State, IdleTime, LogOnTime
        } catch {
            Write-Host "Error querying sessions on ${server}: $_" -ForegroundColor Red
        }
    }

    if (-not $sessions) {
        Write-Host "No active user sessions found on any server." -ForegroundColor Yellow
        return
    }

    # Display sessions in GridView for selection
    $selectedSession = $sessions | Out-GridView -Title "Active Users on NS Servers for AU $AU" -OutputMode Single

    if ($selectedSession) {
        $confirm = Read-Host "Confirm stopping VCA.Sparky.Shell.exe for user $($selectedSession.UserName) on $($selectedSession.Server)? (y/n)"
        if ($confirm -eq 'y') {
            try {
                Invoke-Command -ComputerName $selectedSession.Server -ScriptBlock {
                    param($user)
                    Get-Process -Name "VCA.Sparky.Shell" -IncludeUserName -ErrorAction SilentlyContinue | 
                    Where-Object { $_.UserName -match $user } | 
                    Stop-Process -Force -ErrorAction Stop
                } -ArgumentList $selectedSession.UserName
                Write-Host "VCA.Sparky.Shell.exe stopped for user $($selectedSession.UserName) on $($selectedSession.Server)." -ForegroundColor Green
            } catch {
                Write-Host "Error stopping process: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
        }
    } else {
        Write-Host "No session selected." -ForegroundColor Yellow
    }
}

# Main script logic with menu structure inspired by VCAHospLauncher
$exitScript = $false

while (-not $exitScript) {
    Clear-Host
    Write-Host "Enter the AU number (or 'exit' to quit): " -NoNewline -ForegroundColor Cyan
    $AU = (Read-Host).Trim()

    if ($AU -eq 'exit') {
        $exitScript = $true
        continue
    }

    if ($AU -notmatch '^\d{3,6}$') {
        Write-Host "Invalid AU number. Please enter a 3 to 6 digit number." -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }

    # Display the menu once after entering AU
    Write-Host "`n--- Main Menu for AU $AU ---" -ForegroundColor Green
    Write-Host "0. Change AU" -ForegroundColor White
    Write-Host "1. Abaxis MAC Address Search" -ForegroundColor White
    Write-Host "2. Woofware Errors Check" -ForegroundColor White
    Write-Host "3. Add DHCP Reservation" -ForegroundColor White
    Write-Host "4. User Logon Check" -ForegroundColor White
    Write-Host "5. List AD Users and Check Logon" -ForegroundColor White
    Write-Host "6. Kill Sparky Shell for Logged-in User" -ForegroundColor White
    Write-Host "7. Exit" -ForegroundColor White

    $menuActive = $true
    while ($menuActive) {
        Write-Host "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")][ AU $AU ?]: " -NoNewline -ForegroundColor Cyan
        $choice = (Read-Host).Trim()

        switch ($choice) {
            "0" {
                Write-Host "Returning to AU prompt..." -ForegroundColor Green
                $menuActive = $false
            }
            "1" {
                Abaxis-MacAddressSearch -AU $AU
            }
            "2" {
                Woofware-ErrorsCheck -AU $AU
            }
            "3" {
                Add-DHCPReservation -AU $AU
            }
            "4" {
                User-LogonCheck -AU $AU
            }
            "5" {
                ListADUsersAndCheckLogon -AU $AU
            }
            "6" {
                Kill-SparkyShell -AU $AU
            }
            "7" {
                Write-Host "Exiting..." -ForegroundColor Green
                $exitScript = $true
                $menuActive = $false
            }
            default {
                Write-Host "Invalid choice. Please select 0-7." -ForegroundColor Red
            }
        }
    }
}

# Reset console colors on exit (optional)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "Gray"
Clear-Host