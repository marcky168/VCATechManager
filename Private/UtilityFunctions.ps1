# Consolidated Utility Functions
# Contains: Convert-VcaAU, Copy-ToPSSession, Kill-SparkyShell, Remove-BakRegistry, Update-Changelog

function Convert-VcaAU {
    #Ver. 181211
     #Improved -ilo switch to work correctly with clustered naming convention.
     #Fixed issue with RetainSuffixNumber when a FQDN was provided.
    #Ver. 181115
     #All text is convereted to lowercase to accurately remove duplicates.
     #Added FQDN switch to add domain name to host.
     #Added Domain parameter to be used with -FQDN switch. Defaults to 'vcaantech.com'
     #Added ilo switch to auto-fill suffix.
    #Ver. 181021
     #Added esxihost & db switch.
    #Ver. 181003
     #Added util switch to auto-fill suffix.
     #Added strip switch for extracting AU number.
    #Ver. 180830
     #Added quser switch for quser output.
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('ComputerName', 'Name')]
        [string[]]$AU,
        [string]$Prefix = 'h',
        [string]$Suffix = '-ns',
        [switch]$Clipboard,
        [switch]$NoLeadingZeros,
        [switch]$RetainSuffixNumber,
        [switch]$Quser,
        [switch]$EsxiHost,
        [switch]$Database,
        [switch]$Util,
        [switch]$Ilo,
        [switch]$Strip,
        [switch]$FQDN,
        [string]$Domain = 'vcaantech.com'
    )
    begin {
        # Process clipboard items if clipboard switch was used.
        if ($Clipboard.IsPresent) {
            $AU = Get-Clipboard
        }
        # Remove empty lines
        $AU = $AU | Where-Object { $PSItem }
        # Remove duplicates
        $AU = $AU.ToLower() | Select-Object -Unique

        if ($Database.IsPresent) {
            $Suffix = '-db'
        }
        if ($EsxiHost.IsPresent) {
            $Suffix = '-vm'
        }
        if ($Util.IsPresent) {
            $Suffix = '-util'
        }
        if ($Ilo.IsPresent) {
            $Suffix = '-ilo'
            $RetainSuffixNumber = $true
        }
        if ($FQDN.IsPresent) {
            $Suffix += ".$Domain"
        }
        if ($Strip.IsPresent) {
            $Prefix, $Suffix = ''
            $NoLeadingZeros = $true
        }
    }
    process {
        foreach ($AU_Item in $AU) {
            if ($AU_SuffixNumber) { Clear-Variable -Name AU_SuffixNumber }
            if ($AU_ItemStripped) { Clear-Variable -Name AU_ItemStripped }

            # If input item is numbers only then store in $AU_ItemStripped and skip number extraction.
            if ($AU_Item -match '^[0-9]+$') {
                $AU_ItemStripped = $AU_Item
            }
            # Extract numerical AU # from string starting with case insensitive 'h' or 'au.'
            elseif ($AU_Item -match '^((?i)h|au)[0-9]{2,5}') {
                $AU_ItemStripped = ($AU_Item -replace ('^((?i)h|au)', '') -split '-')[0]

                # Extract suffix count, e.g. -ns01, -ups02
                If ($RetainSuffixNumber.IsPresent) {
                    if ($AU_Item -match '-[a-z]+[0-9]{1,2}') {
                        $AU_SuffixNumber = "$(($AU_Item -split '-')[1] -replace '[^0-9]+')"
                    }
                }
            }
            # Extract numerical AU # from string starting with numbers and leading up to a hyphen.
            elseif ($AU_Item -match '^[0-9]{2,5}-') {
                $AU_ItemStripped = ($AU_Item -split '-')[0]
            }

            # Perform output if AU number format match.
            if ($AU_ItemStripped) {
                # Format output prefix and suffix; fill in leading zeros.
                if ((-not $NoLeadingZeros.IsPresent) -and (-not $Quser.IsPresent)) {
                    if (-not $Ilo.IsPresent) {
                        "$Prefix{0}$AU_ItemStripped$Suffix$AU_SuffixNumber" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                    }
                    else {
                        "$Prefix{0}$AU_ItemStripped-vm$AU_SuffixNumber$Suffix" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                    }
                }
                # Format for quser
                elseif ($Quser.IsPresent) {
                    "quser /server:$Prefix{0}$AU_ItemStripped$Suffix$AU_SuffixNumber" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                }
                # Remove leading zeros
                else {
                    $Prefix + $AU_ItemStripped.TrimStart('0') + $Suffix + $AU_SuffixNumber
                }
            }
            elseif ($AU_Item -like 'hmtprod-*') {
                ($AU_Item -split '-')[0]
            }
        } #foreach
    } #process
    end {
        # Output a return after conversion output when quser switch is specified.
        if ($Quser.IsPresent) { Write-Output "`r" }
    }
} #function

# Harold.Kammermeyer@vca.com
# Copy file via PSSession with path creation and optional hash verification.
function Copy-ToPSSession {
    [CmdletBinding()]
    param(
        [string[]]$ComputerName,
        [string[]]$Path,
        [string]$Destination,
        [switch]$VerifyHash,
        [string]$Algorithm = 'MD5',
        [System.Management.Automation.Runspaces.PSSession[]]$Session,
        [pscredential]$Credential
    )
    begin {
        if (-not (Test-Path -Path $Path)) { Write-Warning "[$Path] Does not exist."; break }

        $FullDestinationPath = "$Destination\$(Split-Path -Path $Path -Leaf)"

        # Create file hash if parameter was used
        if ($VerifyHash.IsPresent) { $Hash = (Get-FileHash -Path $Path -Algorithm $Algorithm).Hash }

        # Create pssession if it doesn't exist
        if (-not $Session) {
            foreach ($ComputerName_Item in $ComputerName) {
                try {
                    $Session = $Session + (New-PSSession -ComputerName $ComputerName_Item -Credential $Credential -ErrorAction Stop)
                }
                catch {
                    Write-Warning "[$ComputerName_Item] $($PSItem.Exception.Message)"
                }
            }
        }
    }
    process {
        foreach ($Session_Item in $Session) {
            # Check if file exists in destination folder
            if (-not (Invoke-Command -Session $Session { Test-Path -Path $using:FullDestinationPath })) {
                # Create destination path if it doesn't exist
                Invoke-Command -Session $Session {
                    if (-not (Test-Path -Path $using:Destination)) { New-Item -ItemType Directory -Path $using:Destination | Out-Null }
                }
                # Copy file to pssession
                try {
                    Write-Host "[$($Session_Item.ComputerName)] Copying $(Split-Path -Path $Path -Leaf)" -ForegroundColor Cyan
                    Copy-Item -Path $Path -Destination $Destination -ToSession $Session_Item -ErrorAction Stop
                }
                catch {
                    Write-Warning $_.Exception.Message
                }
            }
            # Verify hash
            if ($VerifyHash.IsPresent -and $Hash -ne '') {
                # Verify file exists
                $ScriptBlock = {
                    if (Test-Path -Path $using:FullDestinationPath) {
                        if ((Get-FileHash -Path $using:FullDestinationPath -Algorithm $using:Algorithm).Hash -ne $using:Hash) {
                            Write-Warning "Source Hash does not match destination."
                        }
                    }
                }
                Invoke-Command -Session $Session $ScriptBlock
            }
        }
    }
    end {
        if ($ComputerName -and $Session) { Remove-PSSession -Session $Session }
    }
}

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

# Function to update hospital master (moved from main script)
function Update-HospitalMaster {
    param(
        [string]$SharePointBaseUrl = 'https://vca365.sharepoint.com/sites/WOOFconnect/regions',
        [string]$HospitalMasterPath = '/Documents/HOSPITALMASTER.xlsx',
        [string]$DestinationPath = "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx",
        [System.Management.Automation.PSCredential]$ExistingCredential,
        [string]$CredPathSPO
    )

    $CsvPath = "$PSScriptRoot\Private\csv"
    $HospitalMasterXlsx = "$CsvPath\HOSPITALMASTER.xlsx"
    $HospitalMasterXlsxNew = "$CsvPath\HOSPITALMASTER_new.xlsx"

    if (-not (Test-Path -Path $CsvPath)) {
        New-Item -ItemType Directory -Path $CsvPath | Out-Null
    }

    try {
        Write-Status "Connecting to SharePoint Online..." Cyan
        Connect-PnPOnline -Url $SharePointBaseUrl -UseWebLogin -ErrorAction Stop -WarningAction Ignore
        $host.UI.RawUI.BackgroundColor = "Black"
        $host.UI.RawUI.ForegroundColor = "White"

        Write-Status "Downloading HOSPITALMASTER.xlsx..." Cyan
        Get-PnPFile -Url $HospitalMasterPath -Path $CsvPath -Filename 'HOSPITALMASTER_new.xlsx' -AsFile -Force -ErrorAction Stop

        if (Test-Path -Path $HospitalMasterXlsx) {
            $CurrentHash = Get-FileHash -Path $HospitalMasterXlsx -Algorithm SHA256
            $NewHash = Get-FileHash -Path $HospitalMasterXlsxNew -Algorithm SHA256
            if ($CurrentHash.Hash -ne $NewHash.Hash) {
                Write-Status "New version of hospital master found... updating" Green
                Move-Item -Path $HospitalMasterXlsxNew -Destination $HospitalMasterXlsx -Force
                Write-Status "Hospital Master successfully updated" Green
            } else {
                $HospitalFileDate = '{0:M/dd/yyyy h:mm tt}' -f (Get-Item -Path $HospitalMasterXlsx | Select-Object -ExpandProperty LastWriteTime)
                Write-Status "Hospital master XLSX is already up-to-date (Last Write Time: $HospitalFileDate)" Cyan
                Remove-Item -Path $HospitalMasterXlsxNew
            }
        } else {
            Write-Status "Hospital master XLSX downloaded successfully" Green
            Move-Item -Path $HospitalMasterXlsxNew -Destination $HospitalMasterXlsx -Force
        }

        return $true
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Status "Download failed: $($errorMessage)" Red
        if (Test-Path -Path $HospitalMasterXlsxNew) { Remove-Item -Path $HospitalMasterXlsxNew -Force }
        if (Test-Path -Path $HospitalMasterXlsx) {
            $HospitalFileDate = '{0:M/dd/yyyy h:mm tt}' -f (Get-Item -Path $HospitalMasterXlsx | Select-Object -ExpandProperty LastWriteTime)
            Write-Status "Using existing local hospital master XLSX (Last Write Time: $HospitalFileDate)" Yellow
        }
        return $false
    }
}

##############################################################################################
##    Script to delete *.bak profile key from registry / Remove temp profiles from system
##    Author: Lokesh Agarwal
##    Input : servers parameter (Contains Servers name)
##############################################################################################
function Remove-BakRegistry {
	param(
		[string[]]$servers
	)

	Foreach ($server in $servers) {
		##connect with registry of remote machine
		$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Localmachine", "$server")

		##set registry path
		$key = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileList", $true)

		## get all profile name
		$profilereg = $key.GetSubKeyNames()
		$profileregcount = $profilereg.count

		while ($profileregcount -ne 0) {
			## check for bak profiles

			if ($profilereg[$profileregcount - 1] -like "*.bak") {
				$bakname = $profilereg[$profileregcount - 1]

				$baknamefinal = $bakname.Split(".")[0]

				## Delete bak profile
			 $key.DeleteSubKeyTree("$bakname")


				##connect with profileGuid
				$keyGuid = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileGuid", $true)

				## get all profile Guid
				$Guidreg = $keyGuid.GetSubKeyNames()
				$Guidregcount = $Guidreg.count

				while ($Guidregcount -ne 0) {
					$bakname1 = $Guidreg[$Guidregcount - 1]

					$keyGuidTest = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileGuid\$bakname1", $true)
					$KeyGuidSidValue = $keyGuidTest.GetValue("sidstring")
					$KeyGuidSidValue

					if ($baknamefinal -eq $KeyGuidSidValue) {
						## Delete Guid profile
						$keyGuid.DeleteSubKeyTree("$bakname1")
					}
					$Guidregcount = $Guidregcount - 1
				}


			}
			$profileregcount = $profileregcount - 1
		}
	}
} #function

function Update-Changelog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Version,

        [Parameter(Mandatory = $true)]
        [string]$Changes,

        [Parameter(Mandatory = $false)]
        [string]$Date = (Get-Date -Format "yyyy-MM-dd")
    )

    $changelogPath = Join-Path $PSScriptRoot "..\Marc-Tools-Changelog.txt"

    # Read existing changelog if it exists
    $existingContent = if (Test-Path $changelogPath) {
        Get-Content $changelogPath -Raw
    } else {
        ""
    }

    # Create new entry
    $newEntry = "Version $Version - $Date`n$Changes`n`n"

    # Prepend new entry to existing content
    $updatedContent = $newEntry + $existingContent

    # Write back to file
    Set-Content -Path $changelogPath -Value $updatedContent -Encoding UTF8

    Write-Log "Changelog updated for version $Version."
}

# Helper function to normalize MAC addresses by removing hyphens and converting to uppercase
function Normalize-MacAddress {
    param ([string]$MacAddress)
    return $MacAddress.Replace("-", "").Replace(":", "").ToUpper()
}

# Helper function to resolve hostname to IP address with caching
function Resolve-HostIP {
    param (
        [string]$Hostname,
        [hashtable]$IpCache = @{}
    )

    if (-not $IpCache.ContainsKey($Hostname)) {
        try {
            $ipAddresses = [System.Net.Dns]::GetHostAddresses($Hostname)
            $IpCache[$Hostname] = $ipAddresses
        } catch {
            Write-Warning "DNS resolution failed for '$Hostname': $($_.Exception.Message)"
            $IpCache[$Hostname] = $null
        }
    }

    return $IpCache[$Hostname]
}

# Helper function to get DHCP leases for a scope
function Get-DHCPLeases {
    param (
        [string]$DHCPServer,
        [string]$ScopeId
    )

    try {
        $leaseParams = @{
            ComputerName = $DHCPServer
            ScopeId      = $ScopeId
            ErrorAction  = 'Stop'
        }
        return Get-DhcpServerv4Lease @leaseParams
    } catch {
        Write-Warning "Could not retrieve leases from DHCP server '$DHCPServer': $($_.Exception.Message)"
        return $null
    }
}

# Helper function to get DHCP reservations for a scope
function Get-DHCPReservations {
    param (
        [string]$DHCPServer,
        [string]$ScopeId
    )

    try {
        return Get-DhcpServerv4Reservation -ComputerName $DHCPServer -ScopeId $ScopeId -ErrorAction Stop
    } catch {
        Write-Warning "Could not retrieve reservations from DHCP server '$DHCPServer': $($_.Exception.Message)"
        return $null
    }
}

# Helper function to get ARP table from servers
function Get-ARPTable {
    param (
        [string[]]$ComputerNames,
        [pscredential]$Credential
    )

    $arpResults = @()
    $arpJobs = @()

    foreach ($server in $ComputerNames) {
        $job = Start-RSJob -ScriptBlock {
            param($server, $Credential)
            try {
                Invoke-Command -ComputerName $server -ScriptBlock {
                    try {
                        Get-NetNeighbor | Select-Object IPAddress, LinkLayerAddress
                    } catch {
                        # Fallback to arp command if Get-NetNeighbor fails
                        $arpOutput = arp -a
                        $lines = $arpOutput -split "`n" | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+\s+[0-9a-f-]+' }
                        $lines | ForEach-Object {
                            $parts = $_ -split '\s+' | Where-Object { $_ -and $_ -ne 'dynamic' -and $_ -ne 'static' }
                            if ($parts.Count -ge 2) {
                                [PSCustomObject]@{ IPAddress = $parts[0]; LinkLayerAddress = $parts[1] }
                            }
                        }
                    }
                } -Credential $Credential -ErrorAction Stop
            } catch {
                Write-Debug "Failed to get ARP from $server : $($_.Exception.Message)"
                $null
            }
        } -ArgumentList $server, $Credential
        $arpJobs += $job
    }

    $arpResults = $arpJobs | Wait-RSJob | Receive-RSJob | Where-Object { $_ }
    Remove-RSJob -Job $arpJobs

    return $arpResults
}

# Helper function to determine device group based on MAC address
function Get-DeviceGroup {
    param ([string]$MacAddress)

    $macPrefixes = @{
        "VS2"   = @("00-07-32", "00-30-64")
        "HM5"   = @("00-1B-EB")
        "VSPro" = @("00-03-1D")
        "Fuse"  = @("00-90-FB", "00-50-56", "00-0C-29")
    }

    $normalizedMac = Normalize-MacAddress $MacAddress

    foreach ($group in $macPrefixes.Keys) {
        $prefixes = $macPrefixes[$group]
        $normalizedPrefixes = $prefixes | ForEach-Object { Normalize-MacAddress $_ }
        if ($normalizedPrefixes | Where-Object { $normalizedMac.StartsWith($_) }) {
            return $group
        }
    }
    return "Other"
}

# Menu option function: Restart Sparky Services
function Invoke-MenuOption83 {
    param([string]$AU, [pscredential]$ADCredential)

    if (-not (Get-Module -Name ActiveDirectory)) {
        Write-Warning "ActiveDirectory module not available"
        return
    }

    Clear-Variable -Name SiteServers, SiteAU -ErrorAction Ignore
    $SiteAU = Convert-VcaAu -AU $AU -Suffix ''

    # Build Get-ADComputer parameters
    $adParams = @{
        Filter      = "Name -like '$SiteAU-ns*' -and OperatingSystem -like '*Server*'"
        Properties  = 'IPv4Address', 'OperatingSystem'
        Server      = "vcaantech.com"
    }

    # Only add Credential if $ADCredential is not null
    if ($ADCredential) {
        $adParams.Add('Credential', $ADCredential)
    }

    $servers = Get-ADComputer @adParams |
        Select-Object Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object Name

    $selectedServers = $servers | Out-GridView -Title "#83 Select Remote Desktop Server to Reset Sparky Services - v.$((Get-Variable -Name version -ValueOnly)) - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

    if ($selectedServers) {
        Write-Host "Restarting Sparky services on selected servers (parallel processing)..." -ForegroundColor Green
        $restartJobs = @()

        foreach ($server in $selectedServers) {
            $job = Start-RSJob -ScriptBlock {
                param($serverName, $cred)
                try {
                    # Create ScriptBlock dynamically to include server name
                    $remoteScript = [scriptblock]::Create(@"
                        Get-Service -Name Sparky* | Restart-Service -Verbose -ErrorAction Stop
                        "Success: Services restarted on $serverName"
"@)
                    # Build Invoke-Command parameters
                    $invokeParams = @{
                        ComputerName = $serverName
                        ScriptBlock  = $remoteScript
                        ErrorAction  = 'Stop'
                    }

                    # Only add Credential if $cred is not null
                    if ($cred) {
                        $invokeParams.Add('Credential', $cred)
                    }

                    Invoke-Command @invokeParams
                } catch {
                    "Error: Failed to restart services on $serverName - $($_.Exception.Message)"
                }
            } -ArgumentList $server.Name, $ADCredential
            $restartJobs += $job
        }

        # Wait for all jobs to complete and display results
        $results = $restartJobs | Wait-RSJob | Receive-RSJob
        foreach ($result in $results) {
            if ($result -like "Success:*") {
                Write-Host $result -ForegroundColor Green
            } else {
                Write-Host $result -ForegroundColor Red
            }
        }
        Remove-RSJob -Job $restartJobs
    }
}

# Menu option function: ARP Table Viewer and Abaxis MAC Search
function Invoke-MenuOption1 {
    param([string]$AU, [pscredential]$ADCredential, [string]$credPathAD)

    # Prompt for ARP table viewer first to find MAC addresses for filtering
    $wantARP = Read-Host "Do you want to view the ARP table from a server? (y/n)"
    if ($wantARP.ToLower() -eq 'y') {
        # Check credentials
        if (-not $ADCredential -or -not (Test-ADCredentials -Credential $ADCredential)) {
            Write-Host "AD credentials invalid. Prompting for new ones..." -ForegroundColor Yellow
            try {
                $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)" -ErrorAction Stop
                if ($ADCredential) {
                    $ADCredential | Export-Clixml -Path $credPathAD -Force -ErrorAction Stop
                    Write-Host "AD credentials saved." -ForegroundColor Green
                    Write-Log "AD credentials updated via option 1 ARP viewer."
                } else {
                    Write-Host "No credentials provided. Skipping ARP retrieval." -ForegroundColor Yellow
                    return
                }
            } catch {
                Write-Host "Error updating credentials: $($_.Exception.Message)" -ForegroundColor Red
                Write-Log "Error updating credentials in option 1: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
                return
            }
        }

        # Server selection
        Clear-Variable -Name SelectedServers, SiteAU -ErrorAction Ignore
        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
        $adComputerParams = @{
            Filter     = "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'"
            Properties = 'IPv4Address', 'OperatingSystem'
            Server     = "vcaantech.com"
            Credential = $ADCredential
            ErrorAction = 'Stop'
        }

        try {
            $servers = Get-ADComputer @adComputerParams |
                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } |
                Sort-Object -Property Name

            $SelectedServers = $servers | Out-GridView -Title "Select Server(s) to view ARP table" -OutputMode Multiple
        } catch {
            Write-Host "Failed to query servers: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Failed to query servers for ARP viewer AU $AU : $($_.Exception.Message)"
            return
        }

        if (-not $SelectedServers) {
            Write-Host "No servers selected. Skipping ARP retrieval." -ForegroundColor Yellow
        } else {
            # Run remote command
            foreach ($server in $SelectedServers) {
                Write-Host "Pinging IP range on $($server.Name) to populate ARP cache..." -ForegroundColor Green
                try {
                    # Determine which set of credentials to use
                    $params = @{
                        ComputerName = $server.Name
                        ScriptBlock  = {
                            # Get the server's IP to determine subnet
                            $serverIP = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -like '10.*' }).IPAddress | Select-Object -First 1
                            if ($serverIP) {
                                $base = $serverIP -replace '\.\d+$', '.'
                                # Ping sweep ranges: 40-99 and 200-230 using ping.exe for compatibility
                                for ($i = 40; $i -le 99; $i++) {
                                    ping "$base$i" -n 1 -w 100 | Out-Null
                                }
                                for ($i = 200; $i -le 230; $i++) {
                                    ping "$base$i" -n 1 -w 100 | Out-Null
                                }
                            }
                        }
                        ErrorAction  = 'Stop'
                    }

                    # Only add the -Credential parameter if $ADCredential is not null
                    if ($ADCredential) {
                        $params.Add('Credential', $ADCredential)
                    }

                    # Execute the command. It will use the logged-in user if -Credential is not passed.
                    Invoke-Command @params
                } catch {
                    Write-Host "Failed to ping sweep on $($server.Name): $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Ping sweep error for $($server.Name): $($_.Exception.Message)"
                }

                Write-Host "ARP table from $($server.Name):" -ForegroundColor Green
                try {
                    # Determine which set of credentials to use
                    $params = @{
                        ComputerName = $server.Name
                        ScriptBlock  = {
                            try {
                                Get-NetNeighbor | Select-Object IPAddress, LinkLayerAddress | Format-Table -AutoSize
                            } catch {
                                # Fallback to arp command if Get-NetNeighbor fails
                                $arpOutput = arp -a
                                $lines = $arpOutput -split "`n" | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+\s+[0-9a-f-]+' }
                                $lines | ForEach-Object {
                                    $parts = $_ -split '\s+' | Where-Object { $_ -and $_ -ne 'dynamic' -and $_ -ne 'static' }
                                    if ($parts.Count -ge 2) {
                                        [PSCustomObject]@{ IP = $parts[0]; MAC = $parts[1] }
                                    }
                                } | Format-Table -AutoSize
                            }
                        }
                        ErrorAction  = 'Stop'
                    }

                    # Only add the -Credential parameter if $ADCredential is not null
                    if ($ADCredential) {
                        $params.Add('Credential', $ADCredential)
                    }

                    # Execute the command. It will use the logged-in user if -Credential is not passed.
                    Invoke-Command @params
                } catch {
                    Write-Host "Failed to retrieve ARP table from $($server.Name): $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "ARP retrieval error for $($server.Name): $($_.Exception.Message)"
                }
            }
            Write-Host "ARP retrieval complete." -ForegroundColor Green
            Write-Log "Viewed ARP table for servers in AU $AU"
        }
    }

    # Import ActiveDirectory for Abaxis search
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Host "ActiveDirectory module required for Abaxis search. Install RSAT." -ForegroundColor Red
        return
    }

    # Then run the Abaxis MAC Address Search
    Invoke-AbaxisMacSearch -AU $AU
}