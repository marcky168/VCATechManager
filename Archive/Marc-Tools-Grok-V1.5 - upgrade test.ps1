# Combined PowerShell Script with Menu Options

# Set version
$version = "1.5"  # Final version

# Get script path and last write time
$scriptPath = $MyInvocation.MyCommand.Path
if ($scriptPath) {
    $lastWritten = (Get-Item $scriptPath).LastWriteTime.ToString("MM/dd/yyyy HH:mm:ss")
} else {
    $lastWritten = "N/A"
}

# New: Logging toggle and path (create early for initial errors)
$verboseLogging = $false
$logPath = "$PSScriptRoot\logs\marc_tools_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
if (-not (Test-Path "$PSScriptRoot\logs")) { New-Item -Path "$PSScriptRoot\logs" -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message)
    Add-Content -Path $logPath -Value "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")] $Message"
}

# New: Version check at start
try {
    $remoteVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/yourrepo/marc-tools/main/version.txt" -UseBasicParsing -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Content
    if ($remoteVersion -gt $version) {
        Write-Host "New version available: $remoteVersion. Update recommended." -ForegroundColor Yellow
        Write-Log "New version available: $remoteVersion"
    } else {
        Write-Host "Script is up to date (v$version)." -ForegroundColor Green
        Write-Log "Script is up to date (v$version)"
    }
} catch {
    Write-Host "Version check failed. Check internet or URL." -ForegroundColor Yellow
    Write-Log "Version check error: $($_.Exception.Message)"
}

# Import ActiveDirectory module with check
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "ActiveDirectory module failed to load. Install RSAT." -ForegroundColor Red
    Write-Log "ActiveDirectory import error: $($_.Exception.Message)"
}

# Import required modules with try-catch
try {
    Import-Module -Name "$PSScriptRoot\Private\lib\PoshRSJob" -ErrorAction Stop
    Import-Module -Name "$PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction Stop
} catch {
    Write-Host "Module import failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Module import error: $($_.Exception.Message)"
}

# Dot-source functions from Private folder with try-catch
$privateFolder = "$PSScriptRoot\Private"
$privateFiles = Get-ChildItem -Path $privateFolder -Filter *.ps1 -ErrorAction SilentlyContinue
if (-not $privateFiles) {
    Write-Host "No .ps1 files found in $privateFolder. Ensure scripts like ADUserManagement.ps1 are present." -ForegroundColor Red
    Write-Log "No .ps1 files found in $privateFolder."
}
foreach ($file in $privateFiles) {
    try {
        Write-Host "Loading $($file.Name)..." -ForegroundColor Cyan
        . $file.FullName
        Write-Log "Successfully loaded $($file.Name)"
    } catch {
        Write-Host "Failed to dot-source $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Dot-source error for $($file.Name): $($_.Exception.Message)"
    }
}

# Explicitly verify ADUserManagement is loaded
if (-not (Get-Command -Name ADUserManagement -ErrorAction SilentlyContinue)) {
    Write-Host "ADUserManagement function not loaded. Check if ADUserManagement.ps1 exists and is valid in $privateFolder." -ForegroundColor Red
    Write-Log "ADUserManagement function not loaded."
}

# Set window title
$host.UI.RawUI.WindowTitle = "Marc Tools V1 - Grok - $scriptPath (Written:$lastWritten)"

# Set console colors to match the style (dark blue background, white foreground)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# Session cache for valid AUs to reduce AD queries
$validAUs = @{}

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

    Write-Log "Starting Abaxis MAC Address Search for AU $AU"

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

    $hostname = Convert-VcaAu -AU $AU -Suffix '-gw'

    # Resolve hostname to IP address with error handling and retry
    try {
        $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
    } catch {
        Write-Host "DNS resolution failed for '$hostname'. Retrying once..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        try {
            $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
        } catch {
            Write-Host "Error: Could not resolve hostname '$hostname'. $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error in DNS resolution: $($_.Exception.Message)"
            return
        }
    }

    if ($ipAddresses.Length -eq 0) {
        Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
        Write-Log "No IP found for $hostname"
        return
    } elseif ($ipAddresses.Length -gt 1) {
        Write-Host "Warning: Multiple IP addresses found for '$hostname'. Using the first one: $($ipAddresses[0].ToString())" -ForegroundColor Yellow
    }
    $ip = $ipAddresses[0].ToString()

    # Calculate scope ID assuming a /24 subnet (e.g., 192.168.1.0)
    $scopeId = $ip -replace '\.\d+$', '.0'

    # Retrieve DHCP leases for the scope with error handling
    Write-Progress -Activity "Retrieving DHCP leases" -Status "Connecting to $dhcpServer..." -PercentComplete 50
    try {
        $leases = Get-DhcpServerv4Lease -ComputerName $dhcpServer -ScopeId $scopeId -ErrorAction Stop
    } catch {
        Write-Host "Error: Could not retrieve leases from DHCP server '$dhcpServer'. $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Lease retrieval error: $($_.Exception.Message)"
        return
    }

    if (-not $leases) {
        Write-Host "No leases found for scope '$scopeId'."
    }

    # Process each group and find matching leases
    $groupResults = @()
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
            $groupResults += $matchingLeases
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
            $groupResults += $reservations
        } else {
            Write-Host "No reservations found for scope '$scopeId'."
        }
    } catch {
        Write-Host "Error: Could not retrieve reservations from DHCP server '$dhcpServer'. $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Reservation retrieval error: $($_.Exception.Message)"
    }

    # Add nslookup for Hxxxx-fuse
    $fuseHostname = Convert-VcaAu -AU $AU -Suffix '-fuse'
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

                # If ping successful, open the Fuse webpage in Edge new tab (to avoid multiples)
                $fuseUrl = "https://${fuseHostname}:8443"
                Start-Process "msedge" -ArgumentList $fuseUrl
                Write-Host "Opening Fuse webpage: $fuseUrl" -ForegroundColor Green
            } else {
                Write-Host "Fuse device did not respond to ping." -ForegroundColor Red
            }
        } else {
            Write-Host "`nNo IP found for Fuse device ($fuseHostname)." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error resolving Fuse hostname '$fuseHostname' : $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Fuse resolution error: $($_.Exception.Message)"
    }

    Write-Progress -Activity "Retrieving DHCP leases" -Completed

    # New: Export prompt
    $confirmExport = Read-Host "Export results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $groupResults | Export-Csv -Path "$PSScriptRoot\reports\AU$AU_abaxis_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
        Write-Host "Exported to reports folder." -ForegroundColor Green
        Write-Log "Exported Abaxis results for AU $AU"
    }
}

# Function for Woofware Errors Check
function Woofware-ErrorsCheck {
    param([string]$AU)

    Write-Log "Starting Woofware Errors Check for AU $AU"

    try {
        if ($validAUs[$AU]) {
            $servers = $validAUs[$AU]
        } else {
            $servers = Get-VCAServers -AU $AU
            $validAUs[$AU] = $servers
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Log "Error in Woofware check: $($_.Exception.Message)"
        return
    }

    # New: Parallel execution with RSJob
    $jobs = @()
    $totalServers = $servers.Count
    $i = 0
    foreach ($server in $servers) {
        $i++
        Write-Progress -Activity "Querying Woofware errors" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
        $jobs += Start-RSJob -Name $server -ScriptBlock {
            param($server)
            try {
                $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                    $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
                    $serverTime = $using:server + '  ' + $time

                    $allErrors = @()
                    $errors101103 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=101,103} –MaxEvents 50 -ErrorAction SilentlyContinue | Select-Object @{n='Server';e={$using:server}}, @{n='ErrorType';e={"101 or 103"}}, TimeCreated, Id, Message, LevelDisplayName
                    if ($errors101103) { $allErrors += $errors101103 }
                    $errors103 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=103} –MaxEvents 1 -ErrorAction Ignore | Select-Object @{n='Server';e={$using:server}}, @{n='ErrorType';e={"103"}}, TimeCreated, Id, Message, LevelDisplayName
                    if ($errors103) { $allErrors += $errors103 }
                    $errors101 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=101} –MaxEvents 2 -ErrorAction Ignore | Select-Object @{n='Server';e={$using:server}}, @{n='ErrorType';e={"101"}}, TimeCreated, Id, Message, LevelDisplayName
                    if ($errors101) { $allErrors += $errors101 }
                    $errors100 = Get-WinEvent –FilterHashtable @{logname='Application';Providername='Woofware'; level=2 ;id=100} –MaxEvents 10 -ErrorAction Ignore | Select-Object @{n='Server';e={$using:server}}, @{n='ErrorType';e={"100"}}, TimeCreated, Id, Message, LevelDisplayName
                    if ($errors100) { $allErrors += $errors100 }

                    $allErrors
                }
            } catch {
                [PSCustomObject]@{
                    Server = $server
                    ErrorType = "Error"
                    TimeCreated = ""
                    Id = ""
                    Message = "Error : $($_.Exception.Message)"
                    LevelDisplayName = ""
                }
            }
        } -ArgumentList $server
    }

    $results = $jobs | Wait-RSJob | ForEach-Object {
        Receive-RSJob -Job $_
        Remove-RSJob -Job $_
    } | Where-Object { $_ }  # Filter nulls

    Write-Progress -Activity "Querying Woofware errors" -Completed

    # Display results
    $results | Group-Object Server | ForEach-Object {
        Write-Host $_.Name
        $_.Group | Format-Table -AutoSize TimeCreated, ErrorType, Id, LevelDisplayName, Message
    }
    $results | Out-GridView -Title "Woofware Errors for AU $AU"

    # New: Export prompt
    $confirmExport = Read-Host "Export Woofware results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $results | Export-Csv -Path "$PSScriptRoot\reports\AU$AU_woofware_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
        Write-Host "Exported to reports folder." -ForegroundColor Green
        Write-Log "Exported Woofware results for AU $AU"
    }
}

# Function for Add DHCP Reservation
function Add-DHCPReservation {
    param([string]$AU)

    Write-Log "Starting Add DHCP Reservation for AU $AU"

    $macSuffix = Read-Host "Enter MAC suffix after 00-90-FB (e.g., XX-XX-XX)"
    if ($macSuffix -notmatch '^([0-9A-Fa-f]{2}[-:]){2}[0-9A-Fa-f]{2}$') {
        Write-Host "Invalid MAC suffix format. Must be XX-XX-XX." -ForegroundColor Red
        Write-Log "Invalid MAC suffix entered"
        return
    }
    $MACAddress = "00-90-FB-$macSuffix"

    $DHCPServers = "phhospdhcp1.vcaantech.com", "phhospdhcp2.vcaantech.com"
    $ReservationName = "fuse"

    $HospitalNumber = $AU
    if ($HospitalNumber.Length -eq 3) {
        $HospitalNumber = "0" + $HospitalNumber
    }
    $hostname = "h" + $HospitalNumber + "-gw"

    try {
        $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
    } catch {
        Write-Host "Error: Could not resolve hostname '$hostname'. $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Hostname resolution error: $($_.Exception.Message)"
        return
    }

    if ($ipAddresses.Length -eq 0) {
        Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
        Write-Log "No IP for hostname"
        return
    }

    $ip = $ipAddresses[0].IPAddressToString
    $scopeId = $ip -replace '\.[0-9]+$', '.0'
    $scopeId = [System.Net.IPAddress]::Parse($scopeId)
    $ReservationIP = $ip -replace '.[0-9]+$', '.210'

    $results = @()
    $totalServers = $DHCPServers.Count
    $i = 0
    foreach ($Server in $DHCPServers) {
        $i++
        Write-Progress -Activity "Processing DHCP servers" -Status "Server $i of $totalServers : $Server" -PercentComplete (($i / $totalServers) * 100)
        try {
            $ExistingReservation = Get-DhcpServerv4Reservation -ComputerName $Server -IPaddress $ReservationIP -ErrorAction Stop
            if ($ExistingReservation) {
                $Confirm = Read-Host "A DHCP reservation with IP address $ReservationIP and scope $ScopeId already exists on server $Server. Do you want to delete it? (y/n)"
                if ($Confirm.ToLower() -eq "y") {
                    Remove-DhcpServerv4Reservation -ComputerName $Server -IPAddress $ReservationIP -ErrorAction Stop
                    $results += "Deleted DHCP reservation for IP address $ReservationIP and scope $ScopeId on server $Server"
                    Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for $ReservationName"
                    $results += "Added DHCP reservation for IP address $ReservationIP to scope $ScopeId on server $Server"
                }
            } else {
                Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for $ReservationName"
                $results += "Added DHCP reservation for IP address $ReservationIP to scope $ScopeId on server $Server"
            }
        } catch {
            Write-Host "Error with DHCP on $Server : $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "DHCP error on $Server : $($_.Exception.Message)"
            $results += "Error with DHCP on $Server : $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "Processing DHCP servers" -Completed

    $results | Out-String | Write-Host

    # New: Export prompt
    $confirmExport = Read-Host "Export DHCP results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $results | Export-Csv -Path "$PSScriptRoot\reports\AU$AU_dhcp_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
        Write-Host "Exported to reports folder." -ForegroundColor Green
        Write-Log "Exported DHCP results for AU $AU"
    }
}

# Function for User Logon Check
function User-LogonCheck {
    param([string]$AU, [string]$Username)

    Write-Log "Starting User Logon Check for AU $AU, User $Username"

    if (-not $Username) {
        $Username = Read-Host "Enter username"
    }
    if (-not $Username) {
        Write-Host "Username required." -ForegroundColor Red
        return
    }

    try {
        Get-ADUser -Identity $Username -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "User '$Username' not found in AD. Proceed anyway? (y/n)" -ForegroundColor Yellow
        if ((Read-Host).ToLower() -ne 'y') { return }
    }

    try {
        if ($validAUs[$AU]) {
            $servers = $validAUs[$AU]
        } else {
            $servers = Get-VCAServers -AU $AU
            $validAUs[$AU] = $servers
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Log "AU validation error: $($_.Exception.Message)"
        return
    }

    # New: Parallel execution with RSJob
    $jobs = @()
    $totalServers = $servers.Count
    $i = 0
    foreach ($server in $servers) {
        $i++
        Write-Progress -Activity "Querying user logons" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
        $jobs += Start-RSJob -Name $server -ScriptBlock {
            param($server, $Username)
            try {
                $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                    param($Username)
                    # Escape single quotes in Username for XPath
                    $escapedUsername = $Username -replace "'", "''"
                    $filterXPath = "*[System[EventID=4624] and EventData/Data[@Name='TargetUserName']='$escapedUsername']"
                    
                    try {
                        $events = Get-WinEvent -LogName Security -FilterXPath $filterXPath -MaxEvents 5 -ErrorAction Stop | ForEach-Object {
                            $eventXml = [xml]$_.ToXml()
                            $ipAddress = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
                            $_ | Add-Member -NotePropertyName 'IpAddress' -NotePropertyValue $ipAddress -PassThru
                        }
                    } catch {
                        Write-Host "Get-WinEvent error: $($_.Exception.Message)" -ForegroundColor Red
                        $events = @()
                    }

                    if ($events) {
                        $selectedEvent = $events | Where-Object { $_.IpAddress -and $_.IpAddress -ne "-" } | Select-Object -First 1
                        if ($selectedEvent) {
                            [PSCustomObject]@{
                                Server = $env:COMPUTERNAME
                                Username = $Username
                                TimeCreated = $selectedEvent.TimeCreated
                                IpAddress = $selectedEvent.IpAddress
                            }
                        } else {
                            $firstEvent = $events | Select-Object -First 1
                            [PSCustomObject]@{
                                Server = $env:COMPUTERNAME
                                Username = $Username
                                TimeCreated = $firstEvent.TimeCreated
                                IpAddress = "N/A"
                            }
                        }
                    } else {
                        [PSCustomObject]@{
                            Server = $env:COMPUTERNAME
                            Username = $Username
                            TimeCreated = "No logon events found"
                            IpAddress = "N/A"
                        }
                    }
                } -ArgumentList $Username
            } catch {
                [PSCustomObject]@{
                    Server = $server
                    Username = $Username
                    TimeCreated = "Error : Server offline or unreachable - $($_.Exception.Message)"
                    IpAddress = "N/A"
                }
            }
        } -ArgumentList $server, $Username
    }

    $results = $jobs | Wait-RSJob | ForEach-Object {
        Receive-RSJob -Job $_
        Remove-RSJob -Job $_
    }
    Write-Progress -Activity "Querying user logons" -Completed

    # Display results
    $results | Format-Table -AutoSize
    $results | Out-GridView -Title "User Logon Results for AU $AU"

    # New: Export prompt
    $confirmExport = Read-Host "Export logon results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $results | Export-Csv -Path "$PSScriptRoot\reports\AU$AU_logon_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
        Write-Host "Exported to reports folder." -ForegroundColor Green
        Write-Log "Exported logon results for AU $AU"
    }
}

# Function for List AD Users and Check Logon (Option 5)
function ListADUsersAndCheckLogon {
    param([string]$AU)

    Write-Log "Starting List AD Users and Check Logon for AU $AU"

    $groupName = "H" + $AU.PadLeft(4, '0')  # e.g., 'H0966' for AU 966

    try {
        $group = Get-ADGroup -Identity $groupName
        $users = Get-ADGroupMember -Identity $group -Recursive | Get-ADUser -Properties Name, SamAccountName, Department, Title, City, State, telephoneNumber | Select-Object Name, SamAccountName, Department, Title, @{n='Location'; e={$_.City + ', ' + $_.State}}, telephoneNumber
    } catch {
        Write-Host "Error fetching AD group or members: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "AD group fetch error: $($_.Exception.Message)"
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

# Function for Kill Sparky Shell (Option 6)
function Kill-SparkyShell {
    param([string]$AU)

    Write-Log "Starting Kill Sparky Shell for AU $AU"

    try {
        if ($validAUs[$AU]) {
            $servers = $validAUs[$AU]
        } else {
            $servers = Get-VCAServers -AU $AU
            $validAUs[$AU] = $servers
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Log "Server fetch error: $($_.Exception.Message)"
        return
    }

    # New: Parallel for session queries
    $jobs = @()
    $totalServers = $servers.Count
    $i = 0
    foreach ($server in $servers) {
        $i++
        Write-Progress -Activity "Querying active sessions" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
        $jobs += Start-RSJob -Name $server -ScriptBlock {
            param($server)
            try {
                Import-Module -Name "$using:PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue  # Full path in runspace
                Get-TSSession -ComputerName $server -State Active -ErrorAction SilentlyContinue | 
                Where-Object { $_.UserName } | 
                Select-Object @{n='Server'; e={$server}}, UserName, SessionId, State, IdleTime, LogOnTime
            } catch {
                "Error querying sessions on $server : $($_.Exception.Message)"
            }
        } -ArgumentList $server
    }

    $sessions = $jobs | Wait-RSJob | ForEach-Object {
        Receive-RSJob -Job $_
        Remove-RSJob -Job $_
    }
    Write-Progress -Activity "Querying active sessions" -Completed

    if (-not $sessions) {
        Write-Host "No active user sessions found on any server." -ForegroundColor Yellow
        return
    }

    $selectedSession = $sessions | Out-GridView -Title "Active Users on NS Servers for AU $AU" -OutputMode Single

    if ($selectedSession) {
        $confirm = Read-Host "Confirm stopping VCA.Sparky.Shell.exe for user $($selectedSession.UserName) on $($selectedSession.Server)? (y/n)"
        if ($confirm.ToLower() -eq 'y') {
            try {
                Invoke-Command -ComputerName $selectedSession.Server -ScriptBlock {
                    param($user)
                    Get-Process -Name "VCA.Sparky.Shell" -IncludeUserName -ErrorAction SilentlyContinue | 
                    Where-Object { $_.UserName -match $user } | 
                    Stop-Process -Force -ErrorAction Stop
                } -ArgumentList $selectedSession.UserName
                Write-Host "VCA.Sparky.Shell.exe stopped for user $($selectedSession.UserName) on $($selectedSession.Server)." -ForegroundColor Green
                Write-Log "Killed Sparky Shell for user $($selectedSession.UserName) on $($selectedSession.Server)"
            } catch {
                Write-Host "Error stopping process: $($_.Exception.Message)" -ForegroundColor Red
                Write-Log "Error killing Sparky Shell: $($_.Exception.Message)"
            }
        } else {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
        }
    } else {
        Write-Host "No session selected." -ForegroundColor Yellow
    }
}

# Main script logic with menu
$exitScript = $false

while (-not $exitScript) {
    Clear-Host

    # Display tool name and version at top with spacing
    Write-Host "`n`n  Marc Tools V1 - Grok v$version`n" -ForegroundColor Magenta

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

    try {
        if ($validAUs[$AU]) {
            $servers = $validAUs[$AU]
        } else {
            $servers = Get-VCAServers -AU $AU
            $validAUs[$AU] = $servers
        }
    } catch {
        Write-Host "AU $AU invalid or no servers found. Try another?" -ForegroundColor Red
        continue
    }

    # Display the menu once after entering AU
    Write-Host "`n--- Main Menu for AU $AU (v$version) ---" -ForegroundColor Green
    Write-Host "0. Change AU" -ForegroundColor Cyan
    Write-Host "1. Abaxis MAC Address Search" -ForegroundColor Cyan
    Write-Host "2. Woofware Errors Check" -ForegroundColor Cyan
    Write-Host "3. Add DHCP Reservation" -ForegroundColor Cyan
    Write-Host "4. User Logon Check" -ForegroundColor Cyan
    Write-Host "5. List AD Users and Check Logon" -ForegroundColor Cyan
    Write-Host "6. Kill Sparky Shell for Logged-in User" -ForegroundColor Cyan
    Write-Host "7. Exit" -ForegroundColor Cyan
    Write-Host "8. Help" -ForegroundColor Cyan
    Write-Host "9. Toggle Verbose Logging (Current): $(if ($verboseLogging) {'On'} else {'Off'})" -ForegroundColor Cyan
    Write-Host "10. Robo Update" -ForegroundColor Cyan
    Write-Host "11. Update Admin Credentials" -ForegroundColor Cyan
    Write-Host "12. Device Connectivity Test" -ForegroundColor Cyan
    Write-Host "14. AD User Management" -ForegroundColor Cyan

    $menuActive = $true
    while ($menuActive) {
        Write-Host ""
        Write-Host "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")][ AU $AU ?]: " -NoNewline -ForegroundColor Yellow  # Changed color to Yellow
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
            "8" {
                Write-Host "Help Menu:" -ForegroundColor Green
                Write-Host "1. Abaxis MAC Address Search: Searches for Abaxis device MACs in DHCP leases and reservations." -ForegroundColor White
                Write-Host "2. Woofware Errors Check: Checks application logs for Woofware errors on NS servers." -ForegroundColor White
                Write-Host "3. Add DHCP Reservation: Adds or updates DHCP reservation for Fuse device." -ForegroundColor White
                Write-Host "4. User Logon Check: Checks logon events for a user on NS servers." -ForegroundColor White
                Write-Host "5. List AD Users and Check Logon: Lists AD users in hospital group and checks logon for selected." -ForegroundColor White
                Write-Host "6. Kill Sparky Shell: Kills VCA.Sparky.Shell process for selected logged-in user." -ForegroundColor White
                Write-Host "9. Toggle Verbose Logging: Enables/disables logging of actions to file." -ForegroundColor White
                Write-Host "10. Robo Update: Updates script using RoboCopy from network path." -ForegroundColor White
                Write-Host "11. Update Admin Credentials: Update stored admin credentials." -ForegroundColor White
                Write-Host "12. Device Connectivity Test: Test connectivity to devices from DHCP." -ForegroundColor White
                Write-Host "14. AD User Management: Reset password/unlock account for users." -ForegroundColor White
            }
            "9" {
                $verboseLogging = -not $verboseLogging
                Write-Host "Verbose logging now $(if ($verboseLogging) {'enabled'} else {'disabled'})." -ForegroundColor Green
            }
            "10" {
                # Robo Update
                $updateSourcePath = "\\network\path\to\updates"  # Customize
                $confirm = Read-Host "Confirm updating script from $updateSourcePath? (y/n)"
                if ($confirm.ToLower() -eq 'y') {
                    RoboCopy $updateSourcePath $PSScriptRoot * /E /PURGE /R:3 /W:5
                    Write-Host "Update complete. Restarting script..." -ForegroundColor Green
                    & $scriptPath
                    $menuActive = $false
                    $exitScript = $true
                }
            }
            "11" {
    # Update Admin Credentials
    try {
        $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\marcy.admin)"
        if ($newCred) {
            # Store credential using Export-Clixml in script root
            $credPath = "$PSScriptRoot\private\vcaadmin.xml"
            $newCred | Export-Clixml -Path $credPath -Force -ErrorAction Stop
            Write-Host "Admin credentials updated and stored at $credPath." -ForegroundColor Green
            Write-Log "Admin credentials updated for vcaadmin at $credPath."
            # Verify the file was created
            if (Test-Path $credPath) {
                Write-Host "Credential file successfully verified at $credPath." -ForegroundColor Green
            } else {
                Write-Host "Failed to verify credential file at $credPath." -ForegroundColor Red
                Write-Log "Failed to verify credential file at $credPath."
            }
        } else {
            Write-Host "No credentials provided. Operation cancelled." -ForegroundColor Yellow
            Write-Log "Admin credential update cancelled: No credentials provided."
        }
    } catch {
        Write-Host "Error storing credentials: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error storing admin credentials: $($_.Exception.Message)"
    }
}
            "12" {
                # Device Connectivity Test
                DeviceConnectivityTest -AU $AU
            }
            "14" {
                # AD User Management
                ADUserManagement -AU $AU
            }
            default {
                Write-Host "Invalid choice. Please select 0-14." -ForegroundColor Red
            }
        }
    }
}

# Reset console colors on exit (optional)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "Gray"
Clear-Host