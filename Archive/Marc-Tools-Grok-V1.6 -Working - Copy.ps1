# Combined PowerShell Script with Menu Options

# Modification Log:
# 2025-09-19: Enhanced option 5 with parallel Get-TSSession check including ClientIP from event logs; added VNC/Shadow prompts for active users (VNC to client IP via $PSScriptRoot\Private\bin\vncviewer.exe), fallback to event logs. Rationale: Prioritize live actions for better UX and connect to client workstations. Author: Grok. Tested: Isolated snippet passes for sample AU 966.
# 2025-01-01: Updated ListADUsersAndCheckLogon function for parallel active session check with VNC/Shadow prompts, fallback to User-LogonCheck. Added try-catch, Write-Log, Write-Progress. Bumped version to 1.6. Author: Grok.

# Set version
$version = "1.6"  # Updated for enhancements to option 5

# Set console colors to match the style (dark blue background, white foreground) - moved to beginning
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# Load credentials early
$credPathAD = "$PSScriptRoot\Private\vcaadcred.xml"
if (Test-Path $credPathAD) {
    try {
        $ADCredential = Import-Clixml -Path $credPathAD
    } catch {
        $ADCredential = $null
        Write-Host "Failed to load saved AD credentials: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    $ADCredential = $null
}

if (-not $ADCredential) {
    Write-Host "AD Credential not found. Prompting for credentials..." -ForegroundColor Yellow
    $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)"
    if ($ADCredential) {
        $ADCredential | Export-Clixml -Path $credPathAD -Force
        Write-Host "AD credentials saved to $credPathAD." -ForegroundColor Green
        Write-Log "AD credentials saved to $credPathAD."
    } else {
        Write-Host "No AD credentials provided. Some features may not work." -ForegroundColor Yellow
        Write-Log "No AD credentials provided at startup."
    }
}

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
# Create log file early to ensure it exists
New-Item -Path $logPath -ItemType File -Force | Out-Null

# Helper functions moved outside try block for reliability
function Write-Log {
    param([string]$Message)
    Add-Content -Path $logPath -Value "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")] $Message"
}

# Helper function for conditional logging (optimization: only log if verbose)
function Write-ConditionalLog {
    param([string]$Message)
    if ($verboseLogging) {
        Add-Content -Path $logPath -Value "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")] $Message"
    }
}

# Helper function for exporting results (optimization: reduces duplication)
function Export-Results {
    param([array]$Results, [string]$BaseName, [string]$AU)
    $confirmExport = Read-Host "Export results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $exportPath = "$PSScriptRoot\reports\${AU}_${BaseName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Results | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "Exported to $exportPath." -ForegroundColor Green
        Write-Log "Exported $BaseName results for AU $AU"
    }
}

filter Get-PingStatus {
    try {
        $ErrorActionPreference = 'Stop'
        $obj = New-Object system.Net.NetworkInformation.Ping
        if (($obj.Send($PSItem, '1000')).status -eq 'Success') { 'Online!' }
    }
    catch {
        # intentionally left blank
    }
    finally {
        $ErrorActionPreference = 'Continue'
    }
}

# Helper function for cached server fetching (optimization: adds expiration and reduces AD calls)
function Get-CachedServers {
    param($AU)
    $cacheKey = $AU
    $cacheExpiry = 10  # Minutes
    if ($validAUs.ContainsKey($cacheKey) -and $validAUs[$cacheKey].Timestamp -is [DateTime] -and ((Get-Date) - $validAUs[$cacheKey].Timestamp).TotalMinutes -lt $cacheExpiry) {
        return $validAUs[$cacheKey].Servers
    }
    $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
    try {
        $servers = Get-ADComputer -Filter "Name -like '$SiteAU-ns*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Server "vcaantech.com" -Credential $ADCredential | Select-Object -ExpandProperty Name | Sort-Object Name
        $validAUs[$cacheKey] = @{ Servers = $servers; Timestamp = Get-Date }
        return $servers
    } catch {
        Write-Host "Failed to query servers for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Failed to query servers for AU $AU. Error: $($_.Exception.Message)"
        return @()
    }
}

# Function to validate AD credentials
function Test-ADCredentials {
    param([pscredential]$Credential)
    try {
        Get-ADDomain -Credential $Credential -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Set global script root for use in functions
$global:ScriptRoot = $PSScriptRoot

try {
    # Updated: Get-UserSessionsParallel (enhanced for consistency with requirements: added Write-Progress, refined ClientIP fetching, try-catch, logging)
    function Get-UserSessionsParallel {
        param([string]$AU, [string]$Username)
        Write-Log "Starting Get-UserSessionsParallel for AU $AU, User $Username"
        try {
            $servers = Get-CachedServers -AU $AU
        } catch {
            Write-Host "Error fetching servers for AU $AU : $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error in Get-UserSessionsParallel: $($_.Exception.Message)"
            return @()
        }

        $jobs = @()
        $totalServers = $servers.Count
        $i = 0
        foreach ($server in $servers) {
            $i++
            Write-Progress -Activity "Querying user sessions" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server, $Username)
                    try {
                        $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                        Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                            param($Username)
                            Import-Module -Name "$using:PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue
                            $sessions = Get-TSSession -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Active' -or $_.State -eq 'Disconnected' } | Where-Object { $_.UserName -eq $Username }
                            $results = @()
                            foreach ($session in $sessions) {
                                $clientIP = "N/A"
                                try {
                                    $escapedUsername = $Username -replace "'", "''"
                                    $filterXPath = "*[System[EventID=4624] and EventData/Data[@Name='TargetUserName']='$escapedUsername']"
                                    $event = Get-WinEvent -LogName Security -FilterXPath $filterXPath -MaxEvents 1 -ErrorAction Stop | Select-Object -First 1
                                    if ($event) {
                                        $eventXml = [xml]$event.ToXml()
                                        $clientIP = ($eventXml.Event.EventData.Data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
                                        if (-not $clientIP -or $clientIP -eq "-") { $clientIP = "N/A" }
                                    }
                                } catch {
                                    Write-Debug "Failed to fetch ClientIP for session on $env:COMPUTERNAME: $($_.Exception.Message)"
                                }
                                $results += [PSCustomObject]@{
                                    Server     = $env:COMPUTERNAME
                                    UserName   = $session.UserName
                                    SessionId  = $session.SessionId
                                    State      = $session.State
                                    LogOnTime  = $session.LogOnTime
                                    ClientIP   = $clientIP
                                }
                            }
                            $results
                        } -ArgumentList $Username
                    } catch {
                        Write-Debug "Error querying sessions on $server : $($_.Exception.Message)"
                        [PSCustomObject]@{
                            Server     = $server
                            UserName   = $Username
                            SessionId  = "N/A"
                            State      = "Error"
                            LogOnTime  = "N/A"
                            ClientIP   = "N/A"
                        }
                    }
                }
                ArgumentList = $server, $Username
            }
            $jobs += Start-RSJob @jobParams
        }

        $results = $jobs | Wait-RSJob | ForEach-Object { Receive-RSJob -Job $_; Remove-RSJob -Job $_ } | Where-Object { $_ }
        Write-Progress -Activity "Querying user sessions" -Completed
        Write-Log "Get-UserSessionsParallel completed: Found $($results.Count) sessions"
        Write-Debug "Session details: $($results | Out-String)"
        return $results
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
        Write-Host "ActiveDirectory module loaded successfully." -ForegroundColor Cyan  # Debug
    } catch {
        Write-Host "ActiveDirectory module failed to load. Install RSAT." -ForegroundColor Red
        Write-Log "ActiveDirectory import error: $($_.Exception.Message)"
    }

    # Import required modules with try-catch
    try {
        Import-Module -Name "$PSScriptRoot\Private\lib\PoshRSJob" -ErrorAction Stop
        Import-Module -Name "$PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction Stop
        Import-Module -Name "$PSScriptRoot\Private\lib\ImportExcel" -ErrorAction Stop  # Added for HOSPITALMASTER loading
        # Suppress PnP PowerShell update check
        $env:PNPPOWERSHELL_UPDATECHECK = 'Off'
        Import-Module -Name "$PSScriptRoot\Private\lib\PnP.PowerShell" -ErrorAction Stop  # Added for SharePoint access
        Write-Host "Required modules loaded successfully." -ForegroundColor Cyan  # Debug
    } catch {
        Write-Host "Module import failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Module import error: $($_.Exception.Message)"
    }

    # Dot-source functions from Private folder with try-catch
    $privateFolder = "$PSScriptRoot\Private"
    $privateFiles = Get-ChildItem -Path $privateFolder -Filter *.ps1 -ErrorAction SilentlyContinue
    Write-Host "Found private files: $($privateFiles | ForEach-Object { $_.Name })" -ForegroundColor Cyan  # Debug: List files found
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

    # Explicitly verify DeviceConnectivityTest is loaded
    if (-not (Get-Command -Name DeviceConnectivityTest -ErrorAction SilentlyContinue)) {
        Write-Host "DeviceConnectivityTest function not loaded. Check if DeviceConnectivityTest.ps1 exists and is valid in $privateFolder." -ForegroundColor Red
        Write-Log "DeviceConnectivityTest function not loaded."
    }

    # Explicitly verify ListADUsersAndCheckLogon is loaded (added for option 5)
    if (-not (Get-Command -Name ListADUsersAndCheckLogon -ErrorAction SilentlyContinue)) {
        Write-Host "ListADUsersAndCheckLogon function not loaded. Check if ListADUsersAndCheckLogon.ps1 exists and is valid in $privateFolder." -ForegroundColor Red
        Write-Log "ListADUsersAndCheckLogon function not loaded."
    }

    # New: Test log entry
    Write-Log "Script loaded successfully"
    Write-Host "Script loaded successfully. Logging initialized." -ForegroundColor Green  # Debug

    # Session cache for valid AUs to reduce AD queries (updated to use helper)
    $validAUs = @{}

    # Load hospital master to memory (added for hospital info display)
    if (-not $HospitalMaster) {
        $hospitalMasterPath = "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx"
        if (Test-Path $hospitalMasterPath) {
            try {
                $HospitalMaster = Import-Excel -Path $hospitalMasterPath -WorksheetName Misc
                Write-Log "Hospital master loaded successfully from $hospitalMasterPath"
            } catch {
                Write-Host "Failed to load hospital master: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Log "Hospital master load error: $($_.Exception.Message)"
            }
        } else {
            Write-Host "Hospital master file not found at $hospitalMasterPath. Hospital info will not be displayed." -ForegroundColor Yellow
            Write-Log "Hospital master file not found at $hospitalMasterPath"
        }
    }

    # Function to get servers for AU (optimized: better error handling, uses cache helper)
    function Get-VCAServers {
        param([string]$AU)
        if ($AU -notmatch '^\d{3,6}$') {
            throw "Invalid AU number. Please enter a 3 to 6 digit number."
        }
        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
        try {
            $adServers = Get-ADComputer -Filter "Name -like '$SiteAU-ns*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Server "vcaantech.com" -Credential $ADCredential | Select-Object -ExpandProperty Name | Sort-Object Name
            if (-not $adServers) {
                throw "No -NS servers found for AU $AU."
            }
            return $adServers
        } catch {
            if ($_.Exception.Message -like "*credentials*") {
                Write-Host "Warning: AD credentials invalid. Some features may not work. Update via menu option 11." -ForegroundColor Yellow
                Write-Log "AD credentials invalid in Get-VCAServers: $($_.Exception.Message)"
                return @()
            } else {
                Write-Log "Error fetching servers for AU $AU : $($_.Exception.Message)"
                throw
            }
        }
    }

    # New: Function for Abaxis MAC Address Search (optimized: uses helpers, caches IP, splatting)
    function Abaxis-MacAddressSearch {
        param([string]$AU)
        Write-Log "Starting Abaxis MAC Address Search for AU $AU"

        # Cache for IP resolutions (optimization: avoid repeated DNS calls)
        $ipCache = @{}

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

        $dhcpServer = "phhospdhcp2.vcaantech.com"
        $hostname = Convert-VcaAu -AU $AU -Suffix '-gw'

         # Optimized DNS resolution with caching
        if (-not $ipCache.ContainsKey($hostname)) {
            try {
                $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
                $ipCache[$hostname] = $ipAddresses
            } catch {
                Write-Host "DNS resolution failed for '$hostname'. Retrying once..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                try {
                    $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
                    $ipCache[$hostname] = $ipAddresses
                } catch {
                    Write-Host "Error: Could not resolve hostname '$hostname'. $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Error in DNS resolution: $($_.Exception.Message)"
                    return
                }
            }
        } else {
            $ipAddresses = $ipCache[$hostname]
        }

        if ($ipAddresses.Length -eq 0) {
            Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
            Write-Log "No IP found for $hostname"
            return
        } elseif ($ipAddresses.Length -gt 1) {
            Write-Host "Warning: Multiple IP addresses found for '$hostname'. Using the first one: $($ipAddresses[0].ToString())" -ForegroundColor Yellow
        }
        $ip = $ipAddresses[0].ToString()
        $scopeId = $ip -replace '\.\d+$', '.0'

        # Retrieve DHCP leases with splatting (optimization: cleaner code)
        Write-Progress -Activity "Retrieving DHCP leases" -Status "Connecting to $dhcpServer..." -PercentComplete 50
        $leaseParams = @{
            ComputerName = $dhcpServer
            ScopeId      = $scopeId
            ErrorAction  = 'Stop'
        }
        try {
            $leases = Get-DhcpServerv4Lease @leaseParams
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

        # Ping other leased devices (excluding Fuse)
        $otherDevices = $groupResults | Where-Object { $_.ClientId -notmatch '^00-90-FB|^00-50-56|^00-0C-29' }
        $runPingTest = Read-Host "Run ping test on other leased devices? (y/n)"
        if ($runPingTest.ToLower() -eq 'y') {
            foreach ($device in $otherDevices) {
                $ip = $device.IPAddress.ToString()
                $deviceName = if ($device.PSObject.Properties.Match('HostName') -and $device.HostName -and $device.HostName -ne $ip -and $device.HostName -notmatch '^BAD_ADDRESS$') {
                    $device.HostName
                } elseif ($device.PSObject.Properties.Match('Name') -and $device.Name) {
                    $device.Name
                } else {
                    try {
                        $resolved = [System.Net.Dns]::GetHostEntry($ip).HostName
                        if ($resolved -and $resolved -ne $ip) { $resolved } else { $ip }
                    } catch {
                        $ip
                    }
                }
                $pingResult = Test-Connection -ComputerName $ip -Count 2 -Quiet
                Write-Host "Device " -NoNewline
                Write-Host "$deviceName" -ForegroundColor Cyan -NoNewline
                Write-Host " ($ip) : Ping - " -NoNewline
                if ($pingResult) {
                    Write-Host "$pingResult" -ForegroundColor Green
                } else {
                    Write-Host "$pingResult" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Ping test skipped." -ForegroundColor Yellow
        }

       # Add nslookup for Hxxxx-fuse
        $fuseHostname = Convert-VcaAu -AU $AU -Suffix '-fuse'
        if (-not $ipCache.ContainsKey($fuseHostname)) {
            try {
                $fuseIpAddresses = [System.Net.Dns]::GetHostAddresses($fuseHostname)
                $ipCache[$fuseHostname] = $fuseIpAddresses
            } catch {
                Write-Host "`nNo IP found for Fuse device ($fuseHostname)." -ForegroundColor Yellow
                Write-Log "Fuse resolution error: $($_.Exception.Message)"
            }
        } else {
            $fuseIpAddresses = $ipCache[$fuseHostname]
        }

        if ($fuseIpAddresses -and $fuseIpAddresses.Length -gt 0) {
            $fuseIp = $fuseIpAddresses[0].ToString()
            # Determine Fuse type based on IP
            if ($fuseIp -like "10.242*") {
                $fuseType = "Virtual Fuse"
            } else {
                $fuseType = "Physical Fuse"
            }
            Write-Host "`nFuse Device IP ($fuseType from nslookup on $fuseHostname): " -ForegroundColor Green -NoNewline
            Write-Host "$fuseIp" -ForegroundColor Yellow
            $pingResult = Test-Connection -ComputerName $fuseIp -Count 4 -ErrorAction SilentlyContinue
            if ($pingResult) {
                $pingResult | Format-Table -Property Address, ResponseTime, StatusCode
                Write-Host "Fuse device is responsive." -ForegroundColor Green
                $fuseUrl = "https://${fuseHostname}:8443"
                Start-Process "msedge" -ArgumentList $fuseUrl
                Write-Host "Opening Fuse webpage: $fuseUrl" -ForegroundColor Green
            } else {
                Write-Host "Fuse device did not respond to ping." -ForegroundColor Red
                # New: Check if Fuse IP starts with 10.242 and offer vSphere reboot option
                if ($fuseIp -like "10.242*") {
                    $openVSphere = Read-Host "Fuse IP starts with 10.242 and is not responding. Open vSphere to reboot Fuse? (y/n)"
                    if ($openVSphere.ToLower() -eq 'y') {
                        if ($HospitalInfo -and $HospitalInfo.'Time Zone') {
                            $timeZone = $HospitalInfo.'Time Zone'
                            if ($timeZone -in @("America/New_York", "America/Chicago", "America/Detroit", "America/Toronto")) {
                                $vSphereUrl = "https://vcenter.sddc-3-210-64-79.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:93ae2eb5-e9b0-4c7c-b807-ae5f14957305/summary"
                                Write-Host "Launching East/Central Coast vSphere for Fuse reboot." -ForegroundColor Green
                            } elseif ($timeZone -in @("America/Los_Angeles", "America/Denver")) {
                                $vSphereUrl = "https://vcenter.sddc-52-12-159-141.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:7d4e3879-792d-4e6e-85cc-fed91ac7d2c5/summary"
                                Write-Host "Launching West Coast vSphere for Fuse reboot." -ForegroundColor Green
                            } else {
                                Write-Host "Unknown time zone '$timeZone'. Defaulting to East/Central Coast vSphere." -ForegroundColor Yellow
                                $vSphereUrl = "https://vcenter.sddc-3-210-64-79.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:93ae2eb5-e9b0-4c7c-b807-ae5f14957305/summary"
                            }
                            Start-Process $vSphereUrl
                            Write-Log "Launched vSphere for Fuse reboot on AU $AU : $vSphereUrl"
                        } else {
                            Write-Host "Hospital time zone not available. Cannot open vSphere." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "vSphere launch cancelled." -ForegroundColor Yellow
                    }
                }
            }
        }

        Write-Progress -Activity "Retrieving DHCP leases" -Completed

        # Use helper
        Export-Results -Results $groupResults -BaseName "abaxis_results" -AU $AU
    }

    # Function for Woofware Errors Check (optimized: uses helpers, splatting, better job cleanup)
    function Woofware-ErrorsCheck {
        param([string]$AU)
        Write-Log "Starting Woofware Errors Check for AU $AU"

        try {
            $servers = Get-CachedServers -AU $AU
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Log "Error in Woofware check: $($_.Exception.Message)"
            return
        }

        $jobs = @()
        $totalServers = $servers.Count
        $i = 0
        foreach ($server in $servers) {
            $i++
            Write-Progress -Activity "Querying Woofware errors" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server)
                    try {
                        $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                        Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                            $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
                            $serverTime = $using:server + '  ' + $time

                            $allErrors = @()
                            try {
                                $errors100101102 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=100,101,102} -MaxEvents 50 -ErrorAction Stop
                                if ($errors100101102) { $allErrors += $errors100101102 }
                            } catch {
                                Write-Debug "Failed to query Woofware errors 100,101,102 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors102 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=102} -MaxEvents 1 -ErrorAction Stop
                                if ($errors102) { $allErrors += $errors102 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 102 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors101 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=101} -MaxEvents 2 -ErrorAction Stop
                                if ($errors101) { $allErrors += $errors101 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 101 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors100 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=100} -MaxEvents 10 -ErrorAction Stop
                                if ($errors100) { $allErrors += $errors100 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 100 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }

                            $allErrors
                        }
                    } catch {
                        Write-Debug "Error querying sessions on $server : $($_.Exception.Message)"
                        [PSCustomObject]@{
                            Server     = $server
                            UserName   = $Username
                            SessionId  = "N/A"
                            State      = "Error"
                            LogOnTime  = "N/A"
                            ClientIP   = "N/A"
                        }
                    }
                }
                ArgumentList = $server
            }
            $jobs += Start-RSJob @jobParams
        }

        $results = $jobs | Wait-RSJob | ForEach-Object { Receive-RSJob -Job $_; Remove-RSJob -Job $_ } | Where-Object { $_ }
        Write-Progress -Activity "Querying Woofware errors" -Completed

        # Separate results by error ID and display in separate grids
        $errors100 = $results | Where-Object { $_.Id -eq 100 }
        $errors101 = $results | Where-Object { $_.Id -eq 101 }
        $errors102 = $results | Where-Object { $_.Id -eq 102 }

        # Initialize selected error variable
        $selectedError = $null

        # Display errors for ID 100
        if ($errors100) {
            $selected100 = $errors100 | Out-GridView -Title "Woofware Errors ID 100 for AU $AU" -OutputMode Single
            if ($selected100) {
                Write-Host "Selected Error Details (ID 100):" -ForegroundColor Cyan
                $selected100 | Format-List
                $selectedError = $selected100
            }
        }

        # Display errors for ID 101
        if ($errors101) {
            $selected101 = $errors101 | Out-GridView -Title "Woofware Errors ID 101 for AU $AU" -OutputMode Single
            if ($selected101) {
                Write-Host "Selected Error Details (ID 101):" -ForegroundColor Cyan
                $selected101 | Format-List
                $selectedError = $selected101
            }
        }

        # Display errors for ID 102
        if ($errors102) {
            $selected102 = $errors102 | Out-GridView -Title "Woofware Errors ID 102 for AU $AU" -OutputMode Single
            if ($selected102) {
                Write-Host "Selected Error Details (ID 102):" -ForegroundColor Cyan
                $selected102 | Format-List
                $selectedError = $selected102
            }
        }

        # Use helper for export
        Export-Results -Results $results -BaseName "woofware_results" -AU $AU

        # Prompt to send email to dev team
        $sendEmail = Read-Host "Send email to dev team about these errors? (y/n)"
        if ($sendEmail.ToLower() -eq 'y') {
            $description = Read-Host "Enter issue description"

            # Get hospital details from $HospitalInfo (assuming it's in scope; it's loaded earlier per AU)
            if ($HospitalInfo) {
                $location = $HospitalInfo.'Operating Name'
                $contact = $HospitalInfo.'Hospital Manager'
                $phone = $HospitalInfo.'Phone'
            } else {
                $location = "AU $AU"
                $contact = "N/A"
                $phone = "N/A"
            }

            $subject = "AU$($AU.PadLeft(4, '0')) Woofware Error"

            # Build error details string
            if ($selectedError) {
                $errorDetails = @"
Selected Error Details:
Server: $($selectedError.Server)
Error Type: $($selectedError.ErrorType)
Time Created: $($selectedError.TimeCreated)
ID: $($selectedError.Id)
Message: $($selectedError.Message)
Level: $($selectedError.LevelDisplayName)
"@
            } else {
                $errorDetails = "No specific error selected from the grids."
            }

            $recipientChoice = Read-Host "Send to (d)ev team and DBA, or (b) DBA only?"
            switch ($recipientChoice.ToLower()) {
                'd' {
                    $to = "WoofwareDevSupport@vca.com"
                    $cc = "ITSQLDBA@vca.com"
                }
                'b' {
                    $to = "ITSQLDBA@vca.com"
                    $cc = $null
                }
                default {
                    Write-Host "Invalid choice. Defaulting to Dev team and DBA." -ForegroundColor Yellow
                    $to = "WoofwareDevSupport@vca.com"
                    $cc = "ITSQLDBA@vca.com"
                }
            }

            # Read and append default Outlook signature using dummy email method (embed images as base64)
            $signatureHtml = ""
            try {
                # Create Outlook COM object
                $outlook = New-Object -ComObject Outlook.Application

                # Create dummy email to capture default signature
                $dummyMail = $outlook.CreateItem(0)  # 0 = olMailItem
                $dummyMail.Display()  # Briefly shows window to insert signature
                $signatureHtml = $dummyMail.HTMLBody

                # Embed images as base64 to avoid attachment issues
                foreach ($attach in $dummyMail.Attachments) {
                    try {
                        $cid = $attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
                        $data = $attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
                        $base64 = [Convert]::ToBase64String($data)
                        $ext = [System.IO.Path]::GetExtension($attach.FileName).ToLower()
                        $mime = switch ($ext) {
                            '.png' { 'image/png' }
                            '.jpg' { 'image/jpeg' }
                            '.jpeg' { 'image/jpeg' }
                            '.gif' { 'image/gif' }
                            default { 'image/png' }
                        }
                        $signatureHtml = $signatureHtml -replace "cid:$cid", "data:$mime;base64,$base64"
                    } catch {
                        Write-Debug "Failed to embed attachment $($attach.FileName): $($_.Exception.Message)"
                    }
                }

                # Close dummy mail
                $dummyMail.Close(1)  # 1 = olDiscard
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dummyMail) | Out-Null
                $dummyMail = $null  # Prevent re-release in finally
            } catch {
                Write-Host "Failed to load signature via dummy email: $($_.Exception.Message). Proceeding without." -ForegroundColor Yellow
                Write-Log "Signature load error: $($_.Exception.Message)"
            }

            # Build formatted HTML body
            $bodyHtml = @"
<p><strong style="color: red;">Team,</strong></p>
<p><strong>Details:</strong></p>
<p><span style="font-weight: bold; color: red;">Location:</span> <span style="color: #4169e1;">$location</span></p>
<p><span style="font-weight: bold; color: red;">Site Contact Name:</span> <span style="color: #4169e1;">$contact ($phone)</span></p>
<p><span style="font-weight: bold; color: red;">Issue:</span> <span style="color: #4169e1;">$description</span></p>
<p><strong style="color: red;">Error Details:</strong></p>
<pre>$errorDetails</pre>
<br><br>
$signatureHtml
"@

            # Wrap in html/body if signature doesn't include it
            $bodyHtml = "<html><body>$bodyHtml</body></html>"

            try {
                # Create real email
                $mail = $outlook.CreateItem(0)

                # Set email properties
                $mail.To = $to
                if ($cc) { $mail.CC = $cc }
                $mail.Subject = $subject
                $mail.HTMLBody = $bodyHtml  # Set body with embedded signature images

                # Optional: Attach the exported CSV if it was created
                # $exportPath = "$PSScriptRoot\reports\${AU}_woofware_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                # if (Test-Path $exportPath) { $mail.Attachments.Add($exportPath) }

                $mail.Display()  # Opens as draft for review/edit/send
                Write-Host "Email draft created in Outlook with embedded images." -ForegroundColor Green
                Write-Log "Created email draft for Woofware errors AU $AU"

                # Clean up real email and outlook COM objects
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
                $mail = $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                $outlook = $null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            } catch {
                Write-Host "Failed to create email: $($_.Exception.Message). Ensure Outlook is installed." -ForegroundColor Red
                Write-Log "Email creation error: $($_.Exception.Message)"
            } finally {
                # Clean up: Close and release COM objects (only if not already released)
                if ($dummyMail) {
                    $dummyMail.Close(1)  # 1 = olDiscard
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dummyMail) | Out-Null
                    $dummyMail = $null
                }
                if ($mail) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
                    $mail = $null
                }
                if ($outlook) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                    $outlook = $null
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
    }

    # New: Function for Woofware Errors Check by User (variant of Woofware-ErrorsCheck)
    function Woofware-ErrorsCheckByUser {
        param([string]$AU)
        Write-Log "Starting Woofware Errors Check by User for AU $AU"

        # Step 1: Select username from AD group (similar to option 5)
        $adGroupName = 'H' + $AU.PadLeft(4, '0')
        try {
            $groupMembers = Get-ADGroupMember -Identity $adGroupName -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
            $adUsers = $groupMembers | Get-ADUser -Properties Name, SamAccountName -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop | 
                       Select-Object Name, SamAccountName | Sort-Object Name
            if ($adUsers) {
                $selectedUser = $adUsers | Out-GridView -Title "Select user to filter Woofware errors for AU $AU" -OutputMode Single
                if ($selectedUser) {
                    $Username = $selectedUser.SamAccountName
                } else {
                    Write-Host "No user selected. Cancelling." -ForegroundColor Yellow
                    return
                }
            } else {
                Write-Host "No users found in AD group $adGroupName." -ForegroundColor Yellow
                return
            }
        } catch {
            Write-Host "Failed to query AD group '$adGroupName' for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "AD query error in Woofware-ErrorsCheckByUser: $($_.Exception.Message)"
            return
        }

        # Step 2: Query servers (same as original)
        try {
            $servers = Get-CachedServers -AU $AU
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Log "Error in Woofware check by user: $($_.Exception.Message)"
            return
        }

        $jobs = @()
        $totalServers = $servers.Count
        $i = 0
        foreach ($server in $servers) {
            $i++
            Write-Progress -Activity "Querying Woofware errors by user" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server, $Username)
                    try {
                        $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                        Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                            $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
                            $serverTime = $using:server + '  ' + $time

                            $allErrors = @()
                            try {
                                $errors100101102 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=100,101,102} -MaxEvents 50 -ErrorAction Stop
                                if ($errors100101102) { $allErrors += $errors100101102 }
                            } catch {
                                Write-Debug "Failed to query Woofware errors 100,101,102 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors102 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=102} -MaxEvents 1 -ErrorAction Stop
                                if ($errors102) { $allErrors += $errors102 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 102 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors101 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=101} -MaxEvents 2 -ErrorAction Stop
                                if ($errors101) { $allErrors += $errors101 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 101 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }
                            try {
                                $errors100 = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=100} -MaxEvents 10 -ErrorAction Stop
                                if ($errors100) { $allErrors += $errors100 }
                            } catch {
                                Write-Debug "Failed to query Woofware error 100 on $env:COMPUTERNAME: $($_.Exception.Message)"
                            }

                            $allErrors
                        }
                    } catch {
                        Write-Debug "Error querying sessions on $server : $($_.Exception.Message)"
                        [PSCustomObject]@{
                            Server     = $server
                            UserName   = $Username
                            SessionId  = "N/A"
                            State      = "Error"
                            LogOnTime  = "N/A"
                            ClientIP   = "N/A"
                        }
                    }
                }
                ArgumentList = $server, $Username
            }
            $jobs += Start-RSJob @jobParams
        }

        $results = $jobs | Wait-RSJob | ForEach-Object { Receive-RSJob -Job $_; Remove-RSJob -Job $_ } | Where-Object { $_ }
        Write-Progress -Activity "Querying Woofware errors by user" -Completed

        # Now filter by username in the main function
        $filteredErrors = $results | Where-Object { $_.Message -ilike "*VCAANTECH\$Username*" }

        # Combine filtered errors into one collection for single Out-GridView
        $allErrors = $filteredErrors

        # Display all errors in a single grid for selection
        if ($allErrors) {
            $selectedError = $allErrors | Out-GridView -Title "Woofware Errors for User $Username in AU $AU - Select one to email" -OutputMode Single
            if ($selectedError) {
                Write-Host "Selected Error Details:" -ForegroundColor Cyan
                $selectedError | Format-List
            } else {
                Write-Host "No error selected." -ForegroundColor Yellow
            }
        } else {
            Write-Host "No Woofware errors found for user $Username in AU $AU." -ForegroundColor Yellow
            return
        }

        # Export using helper
        Export-Results -Results $results -BaseName "woofware_user_results" -AU $AU

        # Email prompt (same as original, but add username to subject/description)
        $sendEmail = Read-Host "Send email to dev team about these errors? (y/n)"
        if ($sendEmail.ToLower() -eq 'y') {
            $description = Read-Host "Enter issue description"

            # Get hospital details from $HospitalInfo (assuming it's in scope; it's loaded earlier per AU)
            if ($HospitalInfo) {
                $location = $HospitalInfo.'Operating Name'
                $contact = $HospitalInfo.'Hospital Manager'
                $phone = $HospitalInfo.'Phone'
            } else {
                $location = "AU $AU"
                $contact = "N/A"
                $phone = "N/A"
            }

            $subject = "AU$($AU.PadLeft(4, '0')) Woofware Error for User $Username"

            # Build error details string
            if ($selectedError) {
                $errorDetails = @"
Selected Error Details:
Server: $($selectedError.Server)
Error Type: $($selectedError.ErrorType)
Time Created: $($selectedError.TimeCreated)
ID: $($selectedError.Id)
Message: $($selectedError.Message)
Level: $($selectedError.LevelDisplayName)
"@
            } else {
                $errorDetails = "No specific error selected from the grids."
            }

            $recipientChoice = Read-Host "Send to (d)ev team and DBA, or (b) DBA only?"
            switch ($recipientChoice.ToLower()) {
                'd' {
                    $to = "WoofwareDevSupport@vca.com"
                    $cc = "ITSQLDBA@vca.com"
                }
                'b' {
                    $to = "ITSQLDBA@vca.com"
                    $cc = $null
                }
                default {
                    Write-Host "Invalid choice. Defaulting to Dev team and DBA." -ForegroundColor Yellow
                    $to = "WoofwareDevSupport@vca.com"
                    $cc = "ITSQLDBA@vca.com"
                }
            }

            # Read and append default Outlook signature using dummy email method (embed images as base64)
            $signatureHtml = ""
            try {
                # Create Outlook COM object
                $outlook = New-Object -ComObject Outlook.Application

                # Create dummy email to capture default signature
                $dummyMail = $outlook.CreateItem(0)  # 0 = olMailItem
                $dummyMail.Display()  # Briefly shows window to insert signature
                $signatureHtml = $dummyMail.HTMLBody

                # Embed images as base64 to avoid attachment issues
                foreach ($attach in $dummyMail.Attachments) {
                    try {
                        $cid = $attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
                        $data = $attach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102")
                        $base64 = [Convert]::ToBase64String($data)
                        $ext = [System.IO.Path]::GetExtension($attach.FileName).ToLower()
                        $mime = switch ($ext) {
                            '.png' { 'image/png' }
                            '.jpg' { 'image/jpeg' }
                            '.jpeg' { 'image/jpeg' }
                            '.gif' { 'image/gif' }
                            default { 'image/png' }
                        }
                        $signatureHtml = $signatureHtml -replace "cid:$cid", "data:$mime;base64,$base64"
                    } catch {
                        Write-Debug "Failed to embed attachment $($attach.FileName): $($_.Exception.Message)"
                    }
                }

                # Close dummy mail
                $dummyMail.Close(1)  # 1 = olDiscard
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dummyMail) | Out-Null
                $dummyMail = $null  # Prevent re-release in finally
            } catch {
                Write-Host "Failed to load signature via dummy email: $($_.Exception.Message). Proceeding without." -ForegroundColor Yellow
                Write-Log "Signature load error: $($_.Exception.Message)"
            }

            # Build formatted HTML body
            $bodyHtml = @"
<p><strong style="color: #CD5C5C;">Team,</strong></p>
<p><strong>Details:</strong></p>
<p><span style="font-weight: bold; color: #CD5C5C;">Location:</span> <span style="color: #4169e1;">$location</span></p>
<p><span style="font-weight: bold; color: #CD5C5C;">Site Contact Name:</span> <span style="color: #4169e1;">$contact ($phone)</span></p>
<p><span style="font-weight: bold; color: #CD5C5C;">Issue:</span> <span style="color: #4169e1;">$description</span></p>
<p><strong style="color: #CD5C5C;">Error Details:</strong></p>
<pre>$errorDetails</pre>
<br><br>
$signatureHtml
"@

            # Wrap in html/body if signature doesn't include it
            $bodyHtml = "<html><body>$bodyHtml</body></html>"

            try {
                # Create real email
                $mail = $outlook.CreateItem(0)

                # Set email properties
                $mail.To = $to
                if ($cc) { $mail.CC = $cc }
                $mail.Subject = $subject
                $mail.HTMLBody = $bodyHtml  # Set body with embedded signature images

                # Optional: Attach the exported CSV if it was created
                # $exportPath = "$PSScriptRoot\reports\${AU}_woofware_results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                # if (Test-Path $exportPath) { $mail.Attachments.Add($exportPath) }

                $mail.Display()  # Opens as draft for review/edit/send
                Write-Host "Email draft created in Outlook with embedded images." -ForegroundColor Green
                Write-Log "Created email draft for Woofware errors AU $AU"

                # Clean up real email and outlook COM objects
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
                $mail = $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                $outlook = $null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            } catch {
                Write-Host "Failed to create email: $($_.Exception.Message). Ensure Outlook is installed." -ForegroundColor Red
                Write-Log "Email creation error: $($_.Exception.Message)"
            } finally {
                # Clean up: Close and release COM objects (only if not already released)
                if ($dummyMail) {
                    $dummyMail.Close(1)  # 1 = olDiscard
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dummyMail) | Out-Null
                    $dummyMail = $null
                }
                if ($mail) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mail) | Out-Null
                    $mail = $null
                }
                if ($outlook) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                    $outlook = $null
                }
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
    }

    # New: Function for Add DHCP Reservation
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

        $hostname = Convert-VcaAu -AU $AU -Suffix '-gw'

         # Optimized DNS resolution with caching
        if (-not $ipCache.ContainsKey($hostname)) {
            try {
                $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
                $ipCache[$hostname] = $ipAddresses
            } catch {
                Write-Host "DNS resolution failed for '$hostname'. Retrying once..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                try {
                    $ipAddresses = [System.Net.Dns]::GetHostAddresses($hostname)
                    $ipCache[$hostname] = $ipAddresses
                } catch {
                    Write-Host "Error: Could not resolve hostname '$hostname'. $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Error in DNS resolution: $($_.Exception.Message)"
                    return
                }
            }
        } else {
            $ipAddresses = $ipCache[$hostname]
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
            # ...existing code...
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

    # New: Function for Check Active Sessions and Launch (similar to option 77)
    function CheckActiveSessionsAndLaunch {
        param([string]$AU, [array]$Servers)

        Write-Log "Starting Check Active Sessions and Launch for AU $AU"

        # Query sessions using Start-RSJob (similar to whatusers.ps1 in VCAHospLauncher)
        $jobs = @()
        foreach ($server in $Servers) {
            $jobs += Start-RSJob -Name $server -ScriptBlock {
                param($server)
                Import-Module -Name "$using:PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue
                try {
                    Get-TSSession -ComputerName $server -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Active' } | Where-Object { $_.UserName } | Select-Object SessionId, UserName, ClientName, State, IPAddress, IdleTime
                } catch {
                    Write-Warning "Error querying sessions on $server : $($_.Exception.Message)"
                }
            } -ArgumentList $server
        }

        $sessionResults = $jobs | Wait-RSJob | Receive-RSJob
        $jobs | Remove-RSJob

        # Enrich with AD data locally (for Title)
        $enrichedUsers = @()
        foreach ($session in $sessionResults) {
            $adUser = $null
            try {
                $adUser = Get-ADUser -Identity $session.UserName -Properties Title -Credential $ADCredential -ErrorAction SilentlyContinue
            } catch {
                # Skip if user not found in AD
            }
            
            # Handle null/empty values and format IdleTime
            $ipAddress = if ($session.IPAddress) { $session.IPAddress } else { "N/A" }
            $idleTimeFormatted = if ($session.IdleTime) {
                $session.IdleTime.ToString("hh\:mm\:ss")
            } else {
                "00:00:00"
            }
            $title = if ($adUser -and $adUser.Title) { $adUser.Title } else { "N/A" }
            
            $enrichedUsers += [PSCustomObject]@{
                Computer  = $session.PSComputerName  # Use PSComputerName from job
                SessionId = $session.SessionId
                UserName  = $session.UserName
                ClientName = $session.ClientName
                State     = $session.State
                IPAddress = $ipAddress
                IdleTime  = $idleTimeFormatted
                Title     = $title
            }
        }

        if ($enrichedUsers) {
            $SelectedUser = $enrichedUsers | Out-GridView -Title "Active Users on NS Servers for AU $AU - Select to Launch VNC/RDP" -OutputMode Single
            if ($SelectedUser) {
                # Display detailed info like option 77
                $SelectedUser | Format-List
                $launchChoice = Read-Host "Launch VNC (v) or RDP Shadow (r) for $($SelectedUser.UserName) on $($SelectedUser.Computer)? (v/r/n)"
                if ($launchChoice.ToLower() -eq 'v') {
                    # Launch VNC
                    $userIP = $SelectedUser.IPAddress
                    if ($userIP -and $userIP -ne "N/A") {
                        Start-Process "$PSScriptRoot\Private\bin\vncviewer.exe" -ArgumentList "$userIP"
                        Write-Host "Launching VNC for $($SelectedUser.UserName) on $userIP." -ForegroundColor Green
                        Write-Log "Launched VNC for $($SelectedUser.UserName) on $userIP"
                    } else {
                        Write-Host "No IP address available for VNC." -ForegroundColor Red
                    }
                } elseif ($launchChoice.ToLower() -eq 'r') {
                    # Launch RDP Shadow
                    Start-Process "mstsc.exe" -ArgumentList @("/v:$($SelectedUser.Computer)", "/shadow:$($SelectedUser.SessionId)", "/control")
                    Write-Host "Launching RDP Shadow for $($SelectedUser.UserName) on $($SelectedUser.Computer)." -ForegroundColor Green
                    Write-Log "Launched RDP Shadow for $($SelectedUser.UserName) on $($SelectedUser.Computer)"
                } else {
                    Write-Host "Operation cancelled." -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "No active user sessions found on any server." -ForegroundColor Yellow
        }
    }

    # Function for User Logon Check
    function User-LogonCheck {
        param([string]$AU, [string]$Username, [switch]$SkipPropertiesDisplay)

        try {
            Write-Log "Starting User Logon Check for AU $AU, User $Username"

            # List AD users in the hospital group for selection
            $adGroupName = 'H' + $AU.PadLeft(4, '0')
            Write-Log "Debug: User-LogonCheck group name '$adGroupName' for AU $AU"
            try {
                $groupMembers = Get-ADGroupMember -Identity $adGroupName -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
                $adUsers = $groupMembers | Get-ADUser -Properties Name, SamAccountName, EmailAddress, Title -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop | 
                           Select-Object Name, SamAccountName, EmailAddress, Title | Sort-Object Name
                if ($adUsers) {
                    $selectedUser = $adUsers | Out-GridView -Title "Select user from AD group $adGroupName for AU $AU" -OutputMode Single
                    if ($selectedUser) {
                        $Username = $selectedUser.SamAccountName
                    }
                }
            } catch {
                Write-Host "Failed to query AD group '$adGroupName' for AU $AU. Proceeding to manual entry. Error: $($_.Exception.Message)" -ForegroundColor Yellow
            }

            if (-not $Username) {
                $Username = Read-Host "Enter username (or press Enter to cancel)"
            }
            if (-not $Username) {
                Write-Host "Username required." -ForegroundColor Red
                return
            }

            if (-not $SkipPropertiesDisplay) {
                try {
                    # Get domain password policy for expiry calculation
                    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -Server "vcaantech.com" -Credential $ADCredential).MaxPasswordAge
                    $adUser = Get-ADUser -Identity $Username -Properties Name, Title, OfficePhone, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop
                    Write-Host "`nAD Properties for $Username :" -ForegroundColor Cyan
                    $adUser | Select-Object Name, Title, @{n='OfficePhone'; e={$_.OfficePhone}}, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet, @{n='PasswordExpires'; e={ if ($_.PasswordLastSet) { $_.PasswordLastSet + $MaxPasswordAge } else { 'Never Set' } }} | Format-List
                } catch {
                    Write-Host "User '$Username' not found in AD. Proceed anyway? (y/n)" -ForegroundColor Yellow
                    if ((Read-Host).ToLower() -ne 'y') { return }
                }
            }

            try {
                $servers = Get-CachedServers -AU $AU
            } catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
                Write-Log "AU validation error: $($_.Exception.Message)"
                return
            }

            # Check for active sessions first
            $activeSessions = @()
            $totalServers = $servers.Count
            $i = 0
            foreach ($server in $servers) {
                $i++
                Write-Progress -Activity "Checking active sessions for $Username" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
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

            Write-Progress -Activity "Checking active sessions for $Username" -Completed
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
                if ($activeSessions.Count -eq 1) {
                    # If only one session, display details directly in console
                    $selectedSession = $activeSessions[0]
                    Write-Host "Selected Session Details:" -ForegroundColor Cyan
                    Write-Host "User: $($selectedSession.UserName)" -ForegroundColor White
                    Write-Host "Server: $($selectedSession.Server)" -ForegroundColor White
                    Write-Host "Client IP: $($selectedSession.ClientIP)" -ForegroundColor White
                    Write-Host "Client Name: $($selectedSession.ClientName)" -ForegroundColor White
                    if ($selectedSession.ClientIP -eq "N/A") {
                        Write-Host "Note: Client IP is not available. VNC may not work without a valid IP. RDP should work using the server." -ForegroundColor Yellow
                        Write-Host "If VNC fails, manually copy the IP above and connect via your VNC client." -ForegroundColor Yellow
                    } else {
                        Write-Host "Tip: If auto-launch fails, manually copy the IP above for VNC." -ForegroundColor Cyan
                    }
                    Read-Host "Press Enter to continue after reviewing session details"
                    do {
                        $launchChoice = (Read-Host "Launch VNC (v) or RDP Shadow (r) for $($selectedSession.UserName) on $($selectedSession.Server) (IP: $($selectedSession.ClientIP))? (v/r/n)").Trim().ToLower()
                        if ($launchChoice -notin @('v', 'r', 'n')) {
                            Write-Host "Invalid input. Please enter 'v' for VNC, 'r' for RDP Shadow, or 'n' to cancel." -ForegroundColor Yellow
                        }
                    } while ($launchChoice -notin @('v', 'r', 'n'))
                } else {
                    # Multiple sessions: Use Out-GridView for selection, then display details in console
                    $selectedSession = $activeSessions | Out-GridView -Title "Select session for $Username on AU $AU" -OutputMode Single
                    if ($selectedSession) {
                        Write-Host "Selected Session Details:" -ForegroundColor Cyan
                        Write-Host "User: $($selectedSession.UserName)" -ForegroundColor White
                        Write-Host "Server: $($selectedSession.Server)" -ForegroundColor White
                        Write-Host "Client IP: $($selectedSession.ClientIP)" -ForegroundColor White
                        Write-Host "Client Name: $($selectedSession.ClientName)" -ForegroundColor White
                        if ($selectedSession.ClientIP -eq "N/A") {
                            Write-Host "Note: Client IP is not available. VNC may not work without a valid IP. RDP should work using the server." -ForegroundColor Yellow
                            Write-Host "If VNC fails, manually copy the IP above and connect via your VNC client." -ForegroundColor Yellow
                        } else {
                            Write-Host "Tip: If auto-launch fails, manually copy the IP above for VNC." -ForegroundColor Cyan
                        }
                        Read-Host "Press Enter to continue after reviewing session details"
                        do {
                            $launchChoice = (Read-Host "Launch VNC (v) or RDP Shadow (r) for $($selectedSession.UserName) on $($selectedSession.Server) (IP: $($selectedSession.ClientIP))? (v/r/n)").Trim().ToLower()
                            if ($launchChoice -notin @('v', 'r', 'n')) {
                                Write-Host "Invalid input. Please enter 'v' for VNC, 'r' for RDP Shadow, or 'n' to cancel." -ForegroundColor Yellow
                            }
                        } while ($launchChoice -notin @('v', 'r', 'n'))
                    } else {
                        Write-Host "No session selected." -ForegroundColor Yellow
                        return
                    }
                }
                # Launch logic (only if selectedSession is set)
                if ($selectedSession) {
                    if ($launchChoice -eq 'v') {
                        # Launch VNC
                        $vncPath = "$PSScriptRoot\Private\bin\vncviewer.exe"
                        if (Test-Path $vncPath) {
                            $userIP = $selectedSession.ClientIP
                            if ($userIP -and $userIP -ne "N/A" -and $userIP -ne "") {
                                Start-Process $vncPath -ArgumentList $userIP
                                Write-Host "Launching VNC for $($selectedSession.UserName) on $userIP." -ForegroundColor Green
                                Write-Log "Launched VNC for $($selectedSession.UserName) on $userIP"
                            } else {
                                Write-Host "No valid IP address available for VNC. Client IP: '$userIP'. Check DNS or session details." -ForegroundColor Red
                            }
                        } else {
                            Write-Host "VNC viewer not found at $vncPath. Please verify the path." -ForegroundColor Yellow
                        }
                    } elseif ($launchChoice -eq 'r') {
                        # Launch RDP Shadow
                        Start-Process "mstsc.exe" -ArgumentList @("/v:$($selectedSession.Server)", "/shadow:$($selectedSession.SessionId)", "/control")
                        Write-Host "Launching RDP Shadow for $($selectedSession.UserName) on $($selectedSession.Server)." -ForegroundColor Green
                        Write-Log "Launched RDP Shadow for $($selectedSession.UserName) on $($selectedSession.Server)"
                    } else {
                        Write-Host "Operation cancelled." -ForegroundColor Yellow
                    }
                }
            } else {
                # Fall back to event log search if no active sessions
                Write-Host "No active sessions found for $Username. Searching event logs..." -ForegroundColor Yellow
                Write-Log "No active sessions for $Username, proceeding to event log search"

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
                            $invokeParams = @{
                                ComputerName = $server
                                SessionOption = $sessionOption
                                ScriptBlock = {
                                    param($Username)
                                    # Search logon events
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
                                            # Resolve IP to client hostname
                                            $clientName = "N/A"
                                            try {
                                                $clientName = [System.Net.Dns]::GetHostEntry($selectedEvent.IpAddress).HostName
                                            } catch {
                                                $clientName = "N/A"
                                            }
                                            [PSCustomObject]@{
                                                NSServer = $env:COMPUTERNAME
                                                Username = $Username
                                                TimeCreated = $selectedEvent.TimeCreated
                                                IpAddress = $selectedEvent.IpAddress
                                                ClientName = $clientName
                                            }
                                        } else {
                                            $firstEvent = $events | Select-Object -First 1
                                            [PSCustomObject]@{
                                                NSServer = $env:COMPUTERNAME
                                                Username = $Username
                                                TimeCreated = $firstEvent.TimeCreated
                                                IpAddress = "N/A"
                                                ClientName = "N/A"
                                            }
                                        }
                                    } else {
                                        [PSCustomObject]@{
                                            NSServer = $env:COMPUTERNAME
                                            Username = $Username
                                            TimeCreated = "No logon events found"
                                            IpAddress = "N/A"
                                            ClientName = "N/A"
                                        }
                                    }
                                }
                                ArgumentList = $Username
                            }
                            if ($using:ADCredential) {
                                $invokeParams.Credential = $using:ADCredential
                            }
                            Invoke-Command @invokeParams
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
        } catch {
            Write-Host "Error in User-LogonCheck: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error in User-LogonCheck: $($_.Exception.Message)"
        }
    }

    # Function for Kill Sparky Shell (updated to use parallel Invoke-Command for speed)
    function Kill-SparkyShell {
        param([string]$AU)
        Write-Log "Starting Kill Sparky Shell for AU $AU"
        try {
            $servers = Get-CachedServers -AU $AU
            if (-not $servers) {
                Write-Host "No servers found for AU $AU." -ForegroundColor Yellow
                return
            }

            Write-Host "Searching for Sparky processes and files on $($servers.Count) servers..." -ForegroundColor Cyan
            $results = Invoke-Command -ComputerName $servers -ScriptBlock {
                $processes = Get-Process | Where-Object { $_.Name -like "*Sparky*" } -ErrorAction SilentlyContinue
                $results = @()
                foreach ($process in $processes) {
                    try {
                        $owner = (Get-WmiObject -Class Win32_Process -Filter "ProcessId = $($process.Id)").GetOwner()
                        $userName = $owner.User
                    } catch {
                        $userName = "Unknown"
                    }
                    $results += [PSCustomObject]@{
                        Server    = $env:COMPUTERNAME
                        ProcessId = $process.Id
                        ProcessName = $process.Name
                        UserName  = $userName
                        StartTime = $process.StartTime
                    }
                }
                # Search for Sparky files in common locations (limited recursion for performance)
                $files = Get-ChildItem -Path "C:\Program Files", "C:\Program Files (x86)", "C:\" -Filter "*Sparky*" -Recurse -Depth 3 -ErrorAction SilentlyContinue | Select-Object FullName
                [PSCustomObject]@{
                    SparkyResults = $results
                    SparkyFiles = $files
                }
            } -ErrorAction SilentlyContinue

            $processResults = $results | ForEach-Object { $_.SparkyResults } | Where-Object { $_ }
            $sparkyFiles = $results | ForEach-Object { $_.SparkyFiles } | Where-Object { $_ }

            if ($processResults) {
                $selected = $processResults | Out-GridView -Title "Select Sparky process to kill for AU $AU" -OutputMode Single
                if ($selected) {
                    Write-Host "Attempting to kill $($selected.ProcessName) (PID: $($selected.ProcessId)) on $($selected.Server) for user $($selected.UserName)..." -ForegroundColor Yellow
                    Invoke-Command -ComputerName $selected.Server -ScriptBlock {
                        param($processId)
                        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
                    } -ArgumentList $selected.ProcessId -ErrorAction Stop
                    Write-Host "Killed $($selected.ProcessName) process on $($selected.Server)." -ForegroundColor Green
                    Write-Log "Killed $($selected.ProcessName) process ID $($selected.ProcessId) on $($selected.Server) for user $($selected.UserName)"
                } else {
                    Write-Host "No process selected." -ForegroundColor Yellow
                }
            } else {
                Write-Host "No processes with 'Sparky' in the name found on any server for AU $AU." -ForegroundColor Yellow
                if ($sparkyFiles) {
                    Write-Host "Found Sparky-related files on servers (for debugging):" -ForegroundColor Cyan
                    $sparkyFiles | Format-Table -AutoSize
                } else {
                    Write-Host "No Sparky-related files found in common locations on servers." -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Error in Kill-SparkyShell: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error in Kill-SparkyShell: $($_.Exception.Message)"
        }
    }

    # Function for Update Hospital Master (copied from VCAHospLauncher.ps1)
    function Update-HospitalMaster {
        param(
            [pscredential]$EmailCredential
        )
        # Check for Hospital Master update
        if (-not $EmailCredential) { $EmailCredential = Get-StoredCredential -Target vcaemailcreds }

        $HospitalMasterUrl = 'https://vca365.sharepoint.com/sites/WOOFconnect/regions/Documents/HOSPITALMASTER.xlsx'
        $CsvPath = "$PSScriptRoot\Private\csv"
        $HospitalMasterXlsx = "$CsvPath\HOSPITALMASTER.xlsx"
        $HospitalMasterXlsxNew = "$CsvPath\HOSPITALMASTER_new.xlsx"

        if (-not (Test-Path -Path $CsvPath)) { New-Item -ItemType Directory -Path $CsvPath | Out-Null }

        # Download CSV
        try {
            if (Test-Path -Path $HospitalMasterXlsx) {
                #Get-SharePointFile -FileUrl $HospitalMasterUrl -Destination "$PSScriptRoot\private\csv" -NewFileName 'HOSPITALMASTER_new.xlsx' -Credential $EmailCredential -ErrorAction Stop
                try {
                    Get-VcaHospitalMaster -ErrorAction Stop

                    # Get file hash
                    $CurrentHash = Get-FileHash -Path $HospitalMasterXlsx -Algorithm SHA256
                    $NewHash = Get-FileHash -Path $HospitalMasterXlsxNew -Algorithm SHA256

                    # Check if downloaded CSV is newer (different)
                    if ($CurrentHash.Hash -ne $NewHash.Hash) {
                        Move-Item -Path $HospitalMasterXlsxNew -Destination $HospitalMasterXlsx -Force
                        Write-Host "Hospital master updated." -ForegroundColor Green
                    }
                    else {
                        Remove-Item -Path $HospitalMasterXlsxNew -Force
                       
                        Write-Host "Hospital master is up to date." -ForegroundColor Green
                    }
                }
                catch {
                    Write-Warning $PSItem.Exception.Message
                }
            }
            else {
                Write-Host "Hospital master xlsx not found... downloading file`n" -ForegroundColor Cyan
                #Get-SharePointFile -FileUrl $HospitalMasterUrl -Destination "$PSScriptRoot\private\csv" -NewFileName 'HOSPITALMASTER_new.xlsx' -Credential $EmailCredential -ErrorAction Stop
                Get-VcaHospitalMaster -ErrorAction Stop
            }
        }
        catch {
            Write-Warning "[Download failed] $($PSItem.Exception.Message)"
            if (Test-Path -Path $HospitalMasterXlsx) {
                $HospitalFileDate = '{0:M/dd/yyyy h:mm tt}' -f (Get-Item -Path $HospitalMasterXlsx | Select-Object -ExpandProperty LastWriteTime)
                Write-Host "Using local hospital master XLSX (Last Write Time: $HospitalFileDate)`n" -ForegroundColor Cyan
            }
        }
    }

    # Function for Get Vca Hospital Master (copied from VCAHospLauncher.ps1)
    function Get-VcaHospitalMaster {
        [CmdletBinding()]
        param(
            $SharePointUrl = 'https://vca365.sharepoint.com/sites/WOOFconnect/regions'
        )
        try {
            Connect-PnPOnline -Url $SharePointUrl -UseWebLogin -ErrorAction Stop -WarningAction Ignore
            Get-PnPFile -Url '/Documents/HOSPITALMASTER.xlsx' -Path "$PSScriptRoot\private\csv" -Filename 'HOSPITALMASTER_new.xlsx' -AsFile -ErrorAction Stop
        }
        catch {
            throw $_.Exception.Message
        }
    }

    # Function for AD User Management (copied from created file)
    function ADUserManagement {
        param([string]$AU, [pscredential]$Credential)

        $groupName = 'H' + $AU.PadLeft(4, '0')
        Write-Host "Debug: ADUserManagement group name '$groupName' for AU $AU" -ForegroundColor Cyan
        Write-Log "Debug: ADUserManagement group name '$groupName' for AU $AU"

        # Load admin credentials for password reset/unlock
        $adminCredPath = "$PSScriptRoot\Private\vcaadmin.xml"
        if (Test-Path $adminCredPath) {
            try {
                $adminCredential = Import-Clixml -Path $adminCredPath
            } catch {
                Write-Host "Failed to load admin credentials: $($_.Exception.Message)" -ForegroundColor Red
                return
            }
        } else {
            Write-Host "Admin credentials file not found at $adminCredPath. Cannot perform password reset or unlock." -ForegroundColor Red
            return
        }

        try {
            # Get group members and their AD user details
            $groupMembers = Get-ADGroupMember -Identity $groupName -Server "vcaantech.com" -Credential $Credential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
            $users = $groupMembers | Get-ADUser -Properties Name, SamAccountName, EmailAddress, LockedOut, PasswordExpired, LastLogonDate -Server "vcaantech.com" -Credential $Credential -ErrorAction Stop

            if (-not $users) {
                Write-Host "No users found in group $groupName." -ForegroundColor Yellow
                return
            }

            # Display users in grid for selection, sorted by Name
            $selectedUser = $users | Select-Object Name, SamAccountName, EmailAddress, LockedOut, PasswordExpired, LastLogonDate | 
                            Sort-Object Name | Out-GridView -Title "Select user for management in AU $AU" -OutputMode Single

            if ($selectedUser) {
                Write-Host "Selected user: $($selectedUser.Name) ($($selectedUser.SamAccountName))" -ForegroundColor Cyan
                $action = Read-Host "Choose action: (r)eset password, (u)nlock account, (c)ancel"

                switch ($action.ToLower()) {
                    'r' {
                        $newPassword = Read-Host "Enter new password (will be converted to secure string)" -AsSecureString
                        if ($newPassword) {
                            Set-ADAccountPassword -Identity $selectedUser.SamAccountName -NewPassword $newPassword -Credential $adminCredential -ErrorAction Stop
                            Write-Host "Password reset successfully for $($selectedUser.SamAccountName)." -ForegroundColor Green
                        } else {
                            Write-Host "No password entered. Cancelled." -ForegroundColor Yellow
                        }
                    }
                    'u' {
                        Unlock-ADAccount -Identity $selectedUser.SamAccountName -Credential $adminCredential -ErrorAction Stop
                        Write-Host "Account unlocked successfully for $($selectedUser.SamAccountName)." -ForegroundColor Green
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
            }
        } catch {
            Write-Host "Error fetching AD users for group $groupName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Function for Invoke GPUpdate Force (updated to use parallel Invoke-Command for speed)
    function Invoke-GPUpdateForce {
        param([string]$AU)
        Write-Log "Starting GPUpdate /Force for AU $AU"
        try {
            $servers = Get-CachedServers -AU $AU
            if (-not $servers) {
                Write-Host "No servers found for AU $AU." -ForegroundColor Yellow
                return
            }
            $selectedServer = $servers | Out-GridView -Title "Select NS Server for GPUpdate /Force in AU $AU" -OutputMode Single
            if ($selectedServer) {
                Write-Host "Running gpupdate /force on $selectedServer..." -ForegroundColor Cyan
                Invoke-Command -ComputerName $selectedServer -ScriptBlock {
                    gpupdate /force
                } -ErrorAction Stop
                Write-Host "GPUpdate /force completed on $selectedServer." -ForegroundColor Green
                Write-Log "GPUpdate /force completed on $selectedServer for AU $AU"
            } else {
                Write-Host "No server selected." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error running GPUpdate on $selectedServer : $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error in Invoke-GPUpdateForce: $($_.Exception.Message)"
        }
    }

    # Copied and modified from Invoke-DhcpPrompt in VCAHospLauncher
    function Invoke-DhcpScopeSelection {
        [CmdletBinding()]
        param (
            $ComputerName,
            $DhcpServer,
            $DhcpScopes,
            $DhcpScopeId,
            $DhcpScopeName,
            [pscredential]$Credential
        )
        $SitePrefix = $ComputerName  # Use raw AU number instead of Convert-VcaAu to avoid "h" prefix

        try {
            if (-not $DhcpScopes -and -not $DhcpScopeId) {
                Write-Host "`r`n > Invoke-Command -ComputerName $($DhcpServer) { Get-DhcpServerv4Scope | Where-Object { `$_.Name -like `"*AU$SitePrefix-*`" } }" -ForegroundColor Cyan
                $DhcpScopeResults = Invoke-Command -ComputerName $DhcpServer { Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$using:SitePrefix-*" } } -Credential $Credential -ErrorAction Stop |
                    Select-Object -Property ScopeId, SubnetMask, Name, State, StartRange, EndRange, LeaseDuration
                $DhcpScopeResults | Out-GridView -Title 'Select DHCP scope to query' -OutputMode Single -OutVariable DhcpScopeSelection | Out-Null
            } # scope selection
            elseif ($DhcpScopes -and -not $DhcpScopeId) {
                $DhcpScopes | Out-GridView -Title 'Select DHCP scope to query' -OutputMode Single -OutVariable DhcpScopeSelection | Out-Null
            } # refresh scopes
            else {
                $DhcpScopeSelection = {
                    ScopeId   = $DhcpScopeId
                    ScopeName = $DhcpScopeName
                }
            }

            # Return the selected scope instead of showing leases
            if ($DhcpScopeSelection) {
                return $DhcpScopeSelection
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }

    # Main function to select scope and run Angry IP Scanner (modified option 18)
    function Run-AngryIPOnScope {
        param([string]$AU)
        Write-Log "Starting Run Angry IP Scanner on DHCP Scope for AU $AU"

        # Load admin credentials for DHCP server access (required for permissions)
        $adminCredPath = "$PSScriptRoot\Private\vcaadmin.xml"
        if (Test-Path $adminCredPath) {
            try {
                $Credential = Import-Clixml -Path $adminCredPath
                Write-Log "Admin credentials loaded for DHCP access in AU $AU"
            } catch {
                Write-Host "Failed to load admin credentials from $adminCredPath : $($_.Exception.Message). DHCP access may fail." -ForegroundColor Red
                Write-Log "Failed to load admin credentials for DHCP: $($_.Exception.Message)"
                $Credential = $null
            }
        } else {
            Write-Host "Admin credentials file not found at $adminCredPath. Update via menu option 11. DHCP access may fail." -ForegroundColor Yellow
            Write-Log "Admin credentials file missing for DHCP access in AU $AU"
            $Credential = $null
        }

        # Select DHCP server (similar to VCAHospLauncher option 24)
        Write-Host "Selecting DHCP server..." -ForegroundColor Cyan
        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
        $SiteDC = Get-ADComputer -Filter "Name -like '$SiteAU-dc*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address -ErrorAction SilentlyContinue
        $PhoenixDC = Get-ADComputer -Filter "Name -like 'PHHOSPDHCP*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address -ErrorAction SilentlyContinue

        $DhcpServers = @()
        ($SiteDC + $PhoenixDC) | ForEach-Object {
            if ($_) {
                $DhcpServers += [PSCustomObject]@{
                    Name          = $_.Name
                    ADIPv4Address = $_.IPv4Address
                    CanonicalName = $_.CanonicalName
                    Status        = $_.Name | Get-PingStatus
                }
            }
        }

        if (-not $DhcpServers) {
            Write-Host "No DHCP servers found for AU $AU." -ForegroundColor Red
            Write-Log "No DHCP servers found for AU $AU"
            return
        }

        $SelectedDhcpServer = $DhcpServers | Out-GridView -Title "Select DHCP Server for AU $AU" -OutputMode Single
        if (-not $SelectedDhcpServer) {
            Write-Host "No DHCP server selected." -ForegroundColor Red
            Write-Log "No DHCP server selected for AU $AU"
            return
        }

        $DhcpServer = $SelectedDhcpServer.Name
        Write-Host "Selected DHCP Server: $DhcpServer" -ForegroundColor Green

        # Select scope using the function
        $SelectedScope = Invoke-DhcpScopeSelection -ComputerName $AU -DhcpServer $DhcpServer -Credential $Credential

        if ($SelectedScope) {
            $ScopeId = $SelectedScope.ScopeId.ToString()
            $StartIP = $ScopeId -replace '\.0$', '.1'
            $EndIP = $ScopeId -replace '\.0$', '.254'

            Write-Host "Selected Scope: $($SelectedScope.ScopeId) - Name: $($SelectedScope.Name)" -ForegroundColor Yellow
            Write-Host "Running Angry IP Scanner on range: $StartIP - $EndIP" -ForegroundColor Cyan

            # Path to Angry IP Scanner (updated to correct location)
            $AngryIPPath = "$PSScriptRoot\Private\bin\ipscan-win64-3.9.0.exe"

            if (Test-Path $AngryIPPath) {
                # Run Angry IP with range and start scan
                Start-Process -FilePath $AngryIPPath -ArgumentList "-s -f:range $StartIP $EndIP" -WorkingDirectory "$PSScriptRoot\Private\bin"
                Write-Log "Launched Angry IP Scanner for AU $AU on range $StartIP - $EndIP using server $DhcpServer"
            } else {
                Write-Warning "Angry IP Scanner not found at $AngryIPPath. Please install or adjust path."
                Write-Log "Angry IP Scanner not found at $AngryIPPath for AU $AU"
            }
        } else {
            Write-Host "No scope selected." -ForegroundColor Red
            Write-Log "No DHCP scope selected for AU $AU"
        }
    }

# ...existing code...

    # Main script logic with menu
    $exitScript = $false

    while (-not $exitScript) {
        # Reset AU to ensure prompt is shown
        $AU = $null
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

        # Validate AD group for AU
        $adGroupName = 'H' + $AU.PadLeft(4, '0')
        Write-Log "Debug: Validating group name '$adGroupName' for AU $AU"
        try {
            $group = Get-ADObject -Filter "objectClass -eq 'group' -and name -eq '$adGroupName'" -Server "vcaantech.com" -Credential $ADCredential -ErrorAction Stop
            if (-not $group) { throw "Group '$adGroupName' not found." }
        } catch {
            Write-Host "Failed to query AD group '$adGroupName' for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Failed to query AD group '$adGroupName' for AU $AU. Error: $($_.Exception.Message)"
            continue
        }

        # Display hospital information (added, similar to option 14 in VCAHospLauncher)
        if ($HospitalMaster) {
            $HospitalInfo = $HospitalMaster.Where({ $PSItem.'Hospital Number' -eq $AU -or $PSItem.'Hospital Number' -eq "0$AU" -or $PSItem.'Hospital Number' -eq "00$AU" -or $PSItem.'Hospital Number' -eq "000$AU" }) | Select-Object -First 1
            if ($HospitalInfo) {
                Write-Host "`nLocation:" -ForegroundColor Cyan
                Write-Host "$($HospitalInfo.'Operating Name') #$($HospitalInfo.'Hospital Number')"
                Write-Host "$($HospitalInfo.Address)"
                Write-Host "$($HospitalInfo.City), $($HospitalInfo.St) $($HospitalInfo.Zip)"
                Write-Host ''
                Write-Host 'VCA Site Contact:' -ForegroundColor Cyan
                Write-Host "$($HospitalInfo.'Hospital Manager'), $($HospitalInfo.'Hospital Manager Email')"
                Write-Host "$($HospitalInfo.Phone)"

                Write-Host ''
                Write-Host 'Misc. info:' -ForegroundColor Cyan
                Write-Host "Time Zone              : $($HospitalInfo.'Time Zone')"
                Write-Host "URL                    : $($HospitalInfo.GPURL)"
                Write-Host "Back Line              : $($HospitalInfo.'Back Line')"
                Write-Host "System Conversion Date : $($HospitalInfo.'System Conversion Date')"
                Write-Host "System Type            : $($HospitalInfo.'System Type')"
                Write-Host ''


                # Retrieve hospital hours from standard formatted vca site (dynamic fetching)
                Write-Host 'Hours & Info' -ForegroundColor Cyan
                if ($HospitalInfo.GPURL -and ($($HospitalInfo.'Hospital Number') -notmatch '^[8][0-9]{4}$')) {
                    try {
                        $HospitalWeb = Invoke-WebRequest -Uri $HospitalInfo.GPURL -TimeoutSec 10 -ErrorAction Stop
                        if ($HospitalWeb) {
                            $HospitalWebFiltered = $($HospitalWeb.ParsedHtml.body.getElementsByClassName('accordion__item')).innertext
                            if ($HospitalWebFiltered) {
                                $HospitalWebFiltered
                            } else {
                                $HospitalWebFiltered2 = $($HospitalWeb.ParsedHtml.body.getElementsByClassName('hospital-info__middle ')).getElementsByClassName('hospital-info__column col-12 col-md-3 d-flex flex-column')[1].innertext
                                if ($HospitalWebFiltered2) {
                                    $HospitalWebFiltered2
                                }
                            }
                        }
                    } catch {
                        Write-Warning $_.Exception.Message
                        # New: Offer to open hospital website if hours fail to load
                        $openSite = Read-Host "Failed to load hospital hours. Open hospital website in browser? (y/n)"
                        if ($openSite.ToLower() -eq 'y' -and $HospitalInfo.GPURL) {
                            try {
                                Start-Process "msedge" -ArgumentList $HospitalInfo.GPURL
                                Write-Host "Hospital website opened: $($HospitalInfo.GPURL)" -ForegroundColor Green
                            } catch {
                                Write-Host "Failed to open hospital website: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        }
                    }

                    if (-not $HospitalWeb -or (-not $HospitalWebFiltered -and -not $HospitalWebFiltered2)) {
                        try {
                            $HospitalHours = Invoke-RestMethod -Uri "https://uat.vcahospitals.com/api/content/hospital/getUSHospitalHours?HospitalID=$($HospitalInfo.'Hospital Number')" -TimeoutSec 10 -ErrorAction Stop
                            $HospitalHoursFormatted = $HospitalHours | ForEach-Object {
                                [pscustomobject]@{
                                    au                = $_.au
                                    timezone          = $_.timezone
                                    monday_open       = $_.monday_open | Get-Date -Format "h:mm tt"
                                    monday_close      = $_.monday_close | Get-Date -Format "h:mm tt"
                                    monday_all_day    = $_.monday_all_day
                                    tuesday_open      = $_.tuesday_open | Get-Date -Format "h:mm tt"
                                    tuesday_close     = $_.tuesday_close | Get-Date -Format "h:mm tt"
                                    tuesday_all_day   = $_.tuesday_all_day
                                    wednesday_open       = $_.wednesday_open | Get-Date -Format "h:mm tt"
                                    wednesday_close   = $_.wednesday_close | Get-Date -Format "h:mm tt"
                                    wednesday_all_day = $_.wednesday_all_day
                                    thursday_open     = $_.thursday_open | Get-Date -Format "h:mm tt"
                                    thursday_close    = $_.thursday_close | Get-Date -Format "h:mm tt"
                                    thursday_all_day  = $_.thursday_all_day
                                    friday_open       = $_.friday_open | Get-Date -Format "h:mm tt"
                                    friday_close      = $_.friday_close | Get-Date -Format "h:mm tt"
                                    friday_all_day    = $_.friday_all_day
                                    saturday_open     = $_.saturday_open | Get-Date -Format "h:mm tt"
                                    saturday_close    = $_.saturday_close | Get-Date -Format "h:mm tt"
                                    saturday_all_day  = $_.saturday_all_day
                                    sunday_open       = $_.sunday_open | Get-Date -Format "h:mm tt"
                                    sunday_close      = $_.sunday_close | Get-Date -Format "h:mm tt"
                                    sunday_all_day    = $_.sunday_all_day
                                    misc_heading      = $_.misc_heading
                                    misc_hours        = $_.misc_hours -replace '<[^>]+>', ''
                                }
                            }
                            $HospitalHoursFormatted | Format-List
                        } catch {
                            Write-Warning "Failed to fetch hours from API: $($_.Exception.Message)"
                        }
                    }
                } else {
                    # Canada Hours API
                    try {
                        Invoke-RestMethod -Uri "https://uat.vcacanada.com/api/Content/Hospital/GetCAHospitalHours?HospitalID=$($HospitalInfo.'Hospital Number')" -TimeoutSec 10 | Out-String
                    } catch {
                        Write-Warning "Failed to fetch Canada hours: $($_.Exception.Message)"
                    }
                }
                Write-Host ''
                Start-Sleep -Seconds 2  # Pause to ensure info is visible
            } else {
                Write-Host "Hospital AU $AU not found in HOSPITALMASTER.xlsx. Please verify the AU number or update the file." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Hospital master file not loaded. Cannot display hospital information for AU $AU." -ForegroundColor Yellow
        }

        # Calculate hospital time if available
        $hospitalTime = $null
        if ($HospitalInfo -and $HospitalInfo.'Time Zone') {
            try {
                # Map IANA time zones to Windows time zone IDs
                $ianaToWindows = @{
                    "America/Chicago" = "Central Standard Time"
                    "America/New_York" = "Eastern Standard Time"
                    "America/Denver" = "Mountain Standard Time"
                    "America/Los_Angeles" = "Pacific Standard Time"
                    "America/Detroit" = "Eastern Standard Time"
                    "America/Toronto" = "Eastern Standard Time"
                    # Add more mappings
                }
                $timeZoneId = $HospitalInfo.'Time Zone'
                if ($ianaToWindows.ContainsKey($timeZoneId)) {
                    $timeZoneId = $ianaToWindows[$timeZoneId]
                }
                $hospitalTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($timeZoneId)
                $hospitalTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::UtcNow, $hospitalTimeZone).ToString("MM/dd/yyyy h:mm tt")
            } catch {
                $hospitalTime = "N/A"
            }
        } else {
            $hospitalTime = "N/A"
        }

        # Set window title with AU and Timezone
        $host.UI.RawUI.WindowTitle = "Marc Tools V1 - Grok v$version - [$AU] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available'}) - $scriptPath"

        # Display the menu once after entering AU
        Write-Host "`n--- Main Menu for AU $AU (v$version) ---" -ForegroundColor Green
        Write-Host "0. Change AU" -ForegroundColor Cyan
        Write-Host "1. Abaxis MAC Address Search" -ForegroundColor Cyan
        Write-Host "2. Woofware Errors Check" -ForegroundColor Cyan
        Write-Host "2b. Woofware Errors Check by User" -ForegroundColor Cyan
        Write-Host "3. Add DHCP Reservation" -ForegroundColor Cyan
        Write-Host "4. GPUpdate /Force on Selected Server" -ForegroundColor Cyan
        Write-Host "5. List AD Users and Check Logon" -ForegroundColor Cyan
        Write-Host "6. Kill Sparky Shell for Logged-in User" -ForegroundColor Cyan
        Write-Host "7. Exit" -ForegroundColor Cyan
        Write-Host "8. Help" -ForegroundColor Cyan
        Write-Host "9. Toggle Verbose Logging (Current): $(if ($verboseLogging) {'On'} else {'Off'})" -ForegroundColor Cyan
        Write-Host "10. Robo Update" -ForegroundColor Cyan
        Write-Host "11. Update Admin Credentials" -ForegroundColor Cyan
        Write-Host "12. Device Connectivity Test" -ForegroundColor Cyan
        Write-Host "13. Launch ServiceNow for AU Tickets" -ForegroundColor Cyan
        Write-Host "14. AD User Management" -ForegroundColor Cyan
        Write-Host "14u. Update Hospital Master" -ForegroundColor Cyan
        Write-Host "15. Launch vSphere for Fuse VM" -ForegroundColor Cyan
        Write-Host "16. Run Angry IP Scanner on DHCP Scope" -ForegroundColor Cyan
        Write-Host "19. Launch Remote Desktop to Selected Servers" -ForegroundColor Cyan

        $menuActive = $true
        while ($menuActive) {
            # Calculate hospital time dynamically for each prompt
            $hospitalTime = $null
            if ($HospitalInfo -and $HospitalInfo.'Time Zone') {
                try {
                    # Map IANA time zones to Windows time zone IDs
                    $ianaToWindows = @{
                        "America/Chicago" = "Central Standard Time"
                        "America/New_York" = "Eastern Standard Time"
                        "America/Denver" = "Mountain Standard Time"
                        "America/Los_Angeles" = "Pacific Standard Time"

                        "America/Detroit" = "Eastern Standard Time"
                        "America/Toronto" = "Eastern Standard Time"
                        # Add more mappings as needed based on HOSPITALMASTER.xlsx
                    }
                    $timeZoneId = $HospitalInfo.'Time Zone'
                    if ($ianaToWindows.ContainsKey($timeZoneId)) {
                        $timeZoneId = $ianaToWindows[$timeZoneId]
                    }
                    $hospitalTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($timeZoneId)
                    $hospitalTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::UtcNow, $hospitalTimeZone).ToString("MM/dd/yyyy h:mm tt")
                } catch {
                    $hospitalTime = "N/A"
                }
            } else {

                $hospitalTime = "N/A"
            }

            Write-Host ""
            Write-Host "[Local: $(Get-Date -Format "MM/dd/yyyy h:mm tt")] [Hospital: $hospitalTime] [AU $AU ?]: " -NoNewline -ForegroundColor Yellow
            $choice = (Read-Host).Trim()

            switch ($choice) {
                "0" {
                    Clear-Host
                    Write-Host "Returning to AU prompt..." -ForegroundColor Green
                    $menuActive = $false
                    # Reset window title to base
                    $host.UI.RawUI.WindowTitle = "Marc Tools V1 - Grok v$version - $scriptPath"
                }
                "1" {
                    Abaxis-MacAddressSearch -AU $AU
                }
                "2" {
                    Woofware-ErrorsCheck -AU $AU
                }
                "2b" {
                    Woofware-ErrorsCheckByUser -AU $AU
                }
                "3" {
                    Add-DHCPReservation -AU $AU
                }
                "4" {
                    Invoke-GPUpdateForce -AU $AU
                }
                "5" {
                    # Check AD credentials before proceeding
                    if (-not (Test-ADCredentials -Credential $ADCredential)) {
                        Write-Host "AD credentials invalid. Prompting for new ones..." -ForegroundColor Yellow
                        $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)"
                        if ($ADCredential) {
                            $ADCredential | Export-Clixml -Path $credPathAD -Force
                            Write-Host "AD credentials saved." -ForegroundColor Green
                            Write-Log "AD credentials updated via option 5."
                        } else {
                            Write-Host "No credentials provided. Skipping List AD Users and Check Logon." -ForegroundColor Yellow
                            return
                        }
                    }
                    try {
                        User-LogonCheck -AU $AU
                    } catch {
                        Write-Host "Error in option 5: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Log "Error in option 5: $($_.Exception.Message)"
                    }
                }
                "6" {
                    Kill-SparkyShell -AU $AU
                }
                "6d" {
                    # Debug Kill Sparky Shell - list Sparky processes on servers
                    Write-Host "Debug: Listing Sparky processes on servers for AU $AU" -ForegroundColor Cyan
                    Write-Log "Debug: Listing Sparky processes for AU $AU"
                    $servers = Get-CachedServers -AU $AU
                    if (-not $servers) {
                        Write-Host "No servers found for AU $AU." -ForegroundColor Yellow
                        return
                    }
                    foreach ($server in $servers) {
                        Write-Host "Server: $server" -ForegroundColor Yellow
                        try {
                            $procs = Invoke-Command -ComputerName $server -ScriptBlock {
                                Get-Process | Where-Object { $_.Name -like "*Sparky*" } -ErrorAction SilentlyContinue | Select-Object Name, Id, StartTime, @{Name="UserName";Expression={(Get-WmiObject -Class Win32_Process -Filter "ProcessId = $($_.Id)").GetOwner().User}}
                            } -ErrorAction Stop
                            if ($procs) {
                                $procs | Format-Table -AutoSize
                            } else {
                                Write-Host "No Sparky processes found on $server." -ForegroundColor Red
                            }
                        } catch {
                            Write-Host "Error querying $server : $($_.Exception.Message)" -ForegroundColor Red
                            Write-Log "Debug error on $server : $($_.Exception.Message)"
                        }
                    }
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
                    Write-Host "2b. Woofware Errors Check by User: Checks Woofware errors filtered by selected user on NS servers." -ForegroundColor White
                    Write-Host "3. Add DHCP Reservation: Adds or updates DHCP reservation for Fuse device." -ForegroundColor White
                    Write-Host "4. GPUpdate /Force on Selected Server: Runs gpupdate /force on a selected NS server." -ForegroundColor White
                    Write-Host "5. List AD Users and Check Logon: Lists AD users in hospital group and checks logon for selected." -ForegroundColor White
                    Write-Host "6. Kill Sparky Shell: Kills VCA.Sparky.Shell process for selected logged-in user." -ForegroundColor White
                    Write-Host "9. Toggle Verbose Logging: Enables/disables logging of actions to file." -ForegroundColor White
                    Write-Host "10. Robo Update: Updates script using RoboCopy from network path." -ForegroundColor White
                    Write-Host "11. Update Admin Credentials: Update stored admin credentials." -ForegroundColor White
                    Write-Host "12. Device Connectivity Test: Test connectivity to devices from DHCP." -ForegroundColor White
                    Write-Host "13. Launch ServiceNow for AU Tickets" -ForegroundColor White
                    Write-Host "14. AD User Management: Reset password/unlock account for users." -ForegroundColor White
                    Write-Host "14u. Update Hospital Master: Check and download latest HOSPITALMASTER.xlsx from SharePoint." -ForegroundColor White
                    Write-Host "15. Launch vSphere for Fuse VM: Opens vSphere URL for Fuse VM based on hospital location (East/Central or West)." -ForegroundColor White
                    Write-Host "16. Run Angry IP Scanner on DHCP Scope: Selects DHCP scope and runs Angry IP Scanner." -ForegroundColor White
                }
                "9" {
                    $verboseLogging = -not $verboseLogging
                    Write-Host "Verbose logging now $(if ($verboseLogging) {'On'} else {'Off'})." -ForegroundColor Green
               
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
                        $credPath = "$PSScriptRoot\Private\vcaadmin.xml"         # Standardized path with uppercase 'Private'
                        if (Test-Path $credPath) {
                            # Load existing credentials
                            $existingCred = Import-Clixml -Path $credPath -ErrorAction SilentlyContinue
                            if ($existingCred) {
                                Write-Host "Admin credentials are already saved. Current user: $($existingCred.UserName)" -ForegroundColor Green
                                $updateChoice = Read-Host "Do you want to update them? (y/n)"
                                if ($updateChoice.ToLower() -ne 'y') {
                                    Write-Host "Credential update cancelled. Using existing credentials." -ForegroundColor Yellow
                                    Write-Log "Admin credential update cancelled: User chose not to update."
                                } else {
                                    # Prompt for new credentials (if not existing or user chose to update)
                                    $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\marcy.admin)"
                                   
                                    if ($newCred) {
                                        # Store credential using Export-Clixml
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
                                }
                            } else {
                               
                                Write-Host "Existing credential file found but invalid. Re-entering..." -ForegroundColor Yellow
                                # Prompt for new credentials (if not existing or user chose to update)
                                $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\marcy.admin)"
                               
                                if ($newCred) {
                                    # Store credential using Export-Clixml
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
                            }
                        }
                        else {
                            # No existing file, prompt for new credentials
                            $newCred = Get-Credential -Message "Enter new admin credentials (e.g., vcaantech\marcy.admin)"
                           
                            if ($newCred) {
                                # Store credential using Export-Clixml
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
                "13" {
                    # Launch ServiceNow for AU Tickets
                    $snUrl = "https://marsvh.service-now.com/now/nav/ui/classic/params/target/incident_list.do?sysparm_query=u_departmentLIKE$AU&sysparm_first_row=1&sysparm_view="
                    Start-Process $snUrl
                    Write-Host "Opening ServiceNow for AU $AU tickets." -ForegroundColor Green
                }
                "14" {
                    # AD User Management
                    if (-not (Test-ADCredentials -Credential $ADCredential)) {
                        Write-Host "AD credentials invalid. Prompting for new ones..." -ForegroundColor Yellow
                        $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)"
                        if ($ADCredential) {
                            $ADCredential | Export-Clixml -Path $credPathAD -Force
                            Write-Host "AD credentials saved." -ForegroundColor Green
                            Write-Log "AD credentials updated via option 14."
                        } else {
                            Write-Host "No credentials provided. Skipping AD User Management." -ForegroundColor Yellow
                            return
                        }
                    }
                    ADUserManagement -AU $AU -Credential $ADCredential
                }
                "14u" {
                    # Update Hospital Master
                    Update-HospitalMaster
                    # Reload hospital master after update
                    if (Test-Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx") {
                        $HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
                        Write-Host "Hospital master reloaded." -ForegroundColor Green
                    }
                }
                "15" {
                    # Launch vSphere for Fuse VM based on hospital location
                    if ($HospitalInfo -and $HospitalInfo.'Time Zone') {
                        $timeZone = $HospitalInfo.'Time Zone'
                        if ($timeZone -in @("America/New_York", "America/Chicago", "America/Detroit", "America/Toronto")) {
                            $vSphereUrl = "https://vcenter.sddc-3-210-64-79.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:93ae2eb5-e9b0-4c7c-b807-ae5f14957305/summary"
                            Write-Host "Launching East/Central Coast vSphere for AU $AU." -ForegroundColor Green
                        } elseif ($timeZone -in @("America/Los_Angeles", "America/Denver")) {
                            $vSphereUrl = "https://vcenter.sddc-52-12-159-141.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:7d4e3879-792d-4e6e-85cc-fed91ac7d2c5/summary"
                            Write-Host "Launching West Coast vSphere for AU $AU." -ForegroundColor Green
                        } else {
                            Write-Host "Unknown time zone '$timeZone'. Defaulting to East/Central Coast vSphere." -ForegroundColor Yellow
                            $vSphereUrl = "https://vcenter.sddc-3-210-64-79.vmwarevmc.com/ui/app/folder;nav=h/urn:vmomi:Folder:group-d1:93ae2eb5-e9b0-4c7c-b807-ae5f14957305/summary"
                        }
                        Start-Process $vSphereUrl
                        Write-Log "Launched vSphere for AU $AU : $vSphereUrl"
                    } else {
                        Write-Host "Hospital time zone not available. Cannot determine vSphere URL for AU $AU." -ForegroundColor Red
                    }
                }
                "16" {
                    # Run Angry IP Scanner on DHCP Scope
                    Run-AngryIPOnScope -AU $AU
                }
                "19" {
                    # rdc
                    if (Get-Module -Name ActiveDirectory) {
                        Clear-Variable -Name SiteServers, SiteAU -ErrorAction Ignore
                        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
                        Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                            Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                            Out-GridView -Title 'Select Remote Desktop Server(s) to launch' -OutputMode Multiple -OutVariable SiteServers | Out-Null
                        $SiteServers | foreach-object {
                            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$($_.Name) /admin"
                        }
                    }
                    else {
                        if ((-not $Cluster) -and $NetServices) {
                            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$NetServices"
                        }
                        Write-Warning 'ActiveDirectory module not found.'
                        Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                    }
                }
                default {
                    Write-Host "Invalid choice. Please select 0-16." -ForegroundColor Red
                }
            }
        }
    }
} catch {
    Add-Content -Path $logPath -Value "[$(Get-Date -Format "MM/dd/yyyy h:mm tt")] Error during script execution: $($_.Exception.Message)"
    Write-Host "An error occurred during script execution. Check the log file at $logPath for details." -ForegroundColor Red
}

# Reset console colors on exit (optional)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "Gray"
Clear-Host