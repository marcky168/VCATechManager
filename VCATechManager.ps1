# Combined PowerShell Script with Menu Options

# Set version
$version = "1.29"  # MAJOR SECURITY: Configuration abstraction and domain validation

# Load configuration file
$configPath = "$PSScriptRoot\Private\config.json"
$config = $null
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
        Write-Status "Configuration loaded successfully" Green
    } catch {
        Write-Status "Warning: Could not load configuration file. Using default values." Yellow
        Write-Log "Config load error: $($_.Exception.Message)"
    }
} else {
    Write-Status "Warning: Configuration file not found. Using default values." Yellow
    Write-Log "Config file not found at $configPath"
}

# Set default config values if not loaded
if (-not $config) {
    $config = @{
        InternalDomains = @{
            PrimaryDomain = "vcaantech.com"
            SharePointBaseUrl = "https://vca365.sharepoint.com"
            HospitalMasterPath = "/sites/WOOFconnect/regions/Documents/HOSPITALMASTER.xlsx"
        }
        ServerNaming = @{
            DomainPrefix = "h"
            NameServerSuffix = "-ns"
            UtilSuffix = "-util"
            DatabaseSuffix = "-db"
            VMwareSuffix = "-vm"
            ILOSuffix = "-ilo"
        }
        NetworkSettings = @{
            PingTimeout = 100
            MaxConcurrency = 10
            DefaultTimeZone = "Pacific Standard Time"
            DHCPServers = @("phhospdhcp1.vcaantech.com", "phhospdhcp2.vcaantech.com")
            PrimaryDHCPServer = "phhospdhcp2.vcaantech.com"
        }
        SecuritySettings = @{
            RequireDomainJoin = $true
            ValidateCredentials = $true
            CredentialCacheMinutes = 10
        }
    }
}

# Security check: Verify domain membership if required
if ($config.SecuritySettings.RequireDomainJoin) {
    try {
        $domainInfo = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        $isDomainJoined = $domainInfo.PartOfDomain
        $domainName = $domainInfo.Domain

        if (-not $isDomainJoined -or $domainName -notlike "*$($config.InternalDomains.PrimaryDomain)*") {
            Write-Status "SECURITY WARNING: This script requires a domain-joined machine on the $($config.InternalDomains.PrimaryDomain) domain." Red
            Write-Status "Please run this script from a properly configured corporate machine." Red
            Write-Status "Press any key to exit..." Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
    } catch {
        Write-Status "Warning: Could not verify domain membership. Some features may not work." Yellow
        Write-Log "Domain check error: $($_.Exception.Message)"
    }
}

# Set console colors to match the style (dark blue background, white foreground) - moved to beginning
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# Configurable Logging: Default to verbose logging enabled
$verboseLogging = $true

# New: Logging toggle and path (create early for initial errors)
$logPath = "$PSScriptRoot\logs\VCATechManager_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
if (-not (Test-Path "$PSScriptRoot\logs")) { New-Item -Path "$PSScriptRoot\logs" -ItemType Directory -Force | Out-Null }
# Create log file early to ensure it exists
New-Item -Path $logPath -ItemType File -Force | Out-Null

# Helper function for centralized Write-Host customization
function Write-Status ($Message, $Color = "White") {
    Write-Host $Message -ForegroundColor $Color
}

# Helper functions moved outside try block for reliability
function Write-Log {
    param([string]$Message)
    Add-Content -Path $logPath -Value "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] $Message"
}

# Helper function for conditional logging (optimization: only log if verbose)
function Write-ConditionalLog {
    param([string]$Message)
    if ($verboseLogging) {
        Add-Content -Path $logPath -Value "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] $Message"
    }
}

# Function to validate AD credentials with -ErrorAction Stop
function Test-ADCredentials {
    param($Credential)
    if ($Credential -isnot [pscredential]) { return $false }
    try {
        Get-ADDomain -Credential $Credential -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "Test-ADCredentials failed: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
        return $false
    }
}

function Get-ADSecureCredential {
    param(
        [string]$CredPath,
        [string]$PromptMessage = "Enter AD domain credentials (e.g., vcaantech\youruser)"
    )

    # 1. Attempt to Load Saved Credential
    if (Test-Path $CredPath) {
        try {
            $Cred = Import-Clixml -Path $CredPath -ErrorAction Stop
            if ($Cred -is [pscredential] -and (Test-ADCredentials $Cred)) { return $Cred }
        } catch {
            Write-Status "Failed to load saved credentials: $($_.Exception.Message)" Yellow
            Write-Log "Failed to load saved credentials: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
        }
    }

    # 2. Return null if no valid credentials found - let calling code decide whether to prompt
    return $null
}

# Load credentials early with centralized function
$credPathAD = "$PSScriptRoot\Private\vcaadcred.xml"
$ADCredential = Get-ADSecureCredential -CredPath $credPathAD

# Get script path and last write time
$scriptPath = $MyInvocation.MyCommand.Path
if ($scriptPath) {
    $lastWritten = (Get-Item $scriptPath).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
} else {
    $lastWritten = "N/A"
}

# Helper functions moved outside try block for reliability

# Helper function for exporting results (optimization: reduces duplication, added progress bar for large exports)
function Export-Results {
    param([array]$Results, [string]$BaseName, [string]$AU)
    $confirmExport = Read-Host "Export results to CSV? (y/n)"
    if ($confirmExport.ToLower() -eq 'y') {
        $exportPath = "$PSScriptRoot\reports\${AU}_${BaseName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($Results.Count -gt 10) {
            Write-Progress -Activity "Exporting results" -Status "Writing to CSV..." -PercentComplete 50
        }
        $Results | Export-Csv -Path $exportPath -NoTypeInformation
        if ($Results.Count -gt 10) {
            Write-Progress -Activity "Exporting results" -Completed
        }
        Write-Status "Exported to $exportPath." Green
        Write-Log "Exported $BaseName results for AU $AU"
    }
}

# Helper function to ensure Outlook is running for email creation
function Start-OutlookIfNeeded {
    try {
        $outlookProcess = Get-Process -Name outlook -ErrorAction SilentlyContinue
        if (-not $outlookProcess) {
            Write-Status "Starting Outlook..." Yellow
            Start-Process outlook.exe
            Start-Sleep -Seconds 5  # Wait for Outlook to start
        }
    } catch {
        Write-Status "Failed to start Outlook: $($_.Exception.Message)" Red
        Write-Log "Failed to start Outlook: $($_.Exception.Message)"
    }
}

# Helper function for launching VNC with proper error handling
function Start-VNCViewer {
    param([string]$IPAddress, [string]$Username = "", [string]$Computer = "")

    $vncPath = "$PSScriptRoot\Private\bin\vncviewer.exe"

    # Check if VNC executable exists
    if (-not (Test-Path $vncPath)) {
        Write-Status "VNC viewer not found at $vncPath. Please ensure VNC viewer is installed in the correct location." Red
        Write-Status "Expected path: $vncPath" Yellow
        Write-Log "VNC viewer not found at $vncPath"
        return $false
    }

    # Validate IP address
    if (-not $IPAddress -or $IPAddress -eq "N/A" -or $IPAddress -eq "") {
        Write-Status "No valid IP address provided for VNC connection." Red
        Write-Log "Invalid IP address for VNC: '$IPAddress'"
        return $false
    }

    # Validate executable integrity
    try {
        $fileInfo = Get-Item $vncPath -ErrorAction Stop
        if ($fileInfo.Length -lt 1000) {  # Basic check for file size (VNC executables are typically > 1MB)
            throw "VNC executable appears to be corrupted or incomplete (file size: $($fileInfo.Length) bytes)"
        }
    } catch {
        Write-Status "VNC executable validation failed: $($_.Exception.Message)" Red
        Write-Status "Please re-download or re-install the VNC viewer executable." Yellow
        Write-Log "VNC executable validation failed: $($_.Exception.Message)"
        return $false
    }

    # Attempt to launch VNC
    try {
        $argumentList = $IPAddress
        if ($Username) {
            $argumentList += " -UserName=$Username"
        }

        Start-Process -FilePath $vncPath -ArgumentList $argumentList -ErrorAction Stop
        $userInfo = if ($Username) { "$Username on $Computer ($IPAddress)" } else { "$IPAddress" }
        Write-Status "Launching VNC for $userInfo." Green
        Write-Log "Launched VNC for $userInfo"
        return $true
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Status "Failed to launch VNC viewer: $errorMessage" Red
        Write-Log "VNC launch failed: $errorMessage"

        # Provide specific guidance based on error type
        if ($errorMessage -like "*not a valid application*") {
            Write-Status "The VNC executable appears to be corrupted or incompatible with this OS." Yellow
            Write-Status "Please download a fresh copy of the VNC viewer and place it at: $vncPath" Yellow
        } elseif ($errorMessage -like "*access denied*") {
            Write-Status "Access denied to VNC executable. Check file permissions." Yellow
        } elseif ($errorMessage -like "*file not found*") {
            Write-Status "VNC executable not found. Verify the path: $vncPath" Yellow
        } else {
            Write-Status "Unknown error launching VNC. Check the executable and try again." Yellow
        }

        return $false
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
    param($AU, [pscredential]$Credential = $ADCredential)
    $cacheKey = $AU
    $cacheExpiry = 10  # Minutes
    if ($validAUs.ContainsKey($cacheKey) -and $validAUs[$cacheKey].Timestamp -is [DateTime] -and ((Get-Date) - $validAUs[$cacheKey].Timestamp).TotalMinutes -lt $cacheExpiry) {
        return $validAUs[$cacheKey].Servers
    }
    $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
    try {
        # Use splatting for Get-ADComputer to improve readability
        $adComputerParams = @{
            Filter      = "Name -like '$SiteAU-ns*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'"
            Server      = $config.InternalDomains.PrimaryDomain
            ErrorAction = 'Stop'
        }

        # Only add Credential if provided
        if ($Credential) {
            $adComputerParams.Add('Credential', $Credential)
        }

        $servers = Get-ADComputer @adComputerParams | Select-Object -ExpandProperty Name | Sort-Object Name
        $validAUs[$cacheKey] = @{ Servers = $servers; Timestamp = Get-Date }
        return $servers
    } catch {
        Write-Host "Failed to query servers for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Failed to query servers for AU $AU. Error: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
        return @()
    }
}

# Function to make API call with optional auth and retries
function Invoke-GitHubApi {
    param($url, $headers, $pat, $patPath, $retries = 3)
    for ($i = 1; $i -le $retries; $i++) {
        try {
            Write-Host "Attempting API call (try $i/$retries): $url" -ForegroundColor Cyan
            $response = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing -ErrorAction Stop
            return $response
        } catch {
            $statusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.Value__ } else { $null }
            Write-Host "API call failed: $url - Status: $statusCode - $($_.Exception.Message)" -ForegroundColor Red
            if ($statusCode -eq 404) {
                if (-not $pat) {
                    Write-Host "404 detected�repo may be private. Enter GitHub PAT (leave blank if public):" -ForegroundColor Yellow
                    $patInput = Read-Host
                    if ($patInput) {
                        $securePat = ConvertTo-SecureString $patInput -AsPlainText -Force
                        $pat = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePat))
                        $pat | Set-Content $patPath
                        Write-Host "PAT saved for future use." -ForegroundColor Green
                    }
                }
                if ($pat) {
                    $authHeaders = $headers.Clone()
                    $authHeaders["Authorization"] = "Bearer $pat"
                    Write-Host "Retrying with auth..." -ForegroundColor Cyan
                    $response = Invoke-WebRequest -Uri $url -Headers $authHeaders -UseBasicParsing -ErrorAction Stop
                    return $response
                } else {
                    throw "No PAT provided for private repo."
                }
            } elseif ($i -lt $retries) {
                Write-Host "Retrying in 5 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 5
            } else {
                throw
            }
        }
    }
}

function Sync-Repo {
    $owner = "marcky168"
    $repo = "VCATechManager"
    $branch = "main"
    $lastCommitShaFile = "$PSScriptRoot\last_commit_sha.txt"
    $apiHeaders = @{
        Accept = "application/vnd.github+json"
        "User-Agent" = "VCATechManager-Script/$version"
    }

    $patPath = "$PSScriptRoot\Private\github_pat.txt"
    $pat = $null
    if (Test-Path $patPath) {
        $pat = Get-Content $patPath -Raw
    }

    # Get latest commit
    $commitUrl = "https://api.github.com/repos/$owner/$repo/commits/$branch"
    $commitResponse = Invoke-GitHubApi -url $commitUrl -headers $apiHeaders -pat $pat -patPath $patPath
    $commitData = ConvertFrom-Json $commitResponse.Content
    $remoteCommitSha = $commitData.sha.Trim()
    Write-Host "Remote commit SHA: $remoteCommitSha" -ForegroundColor Green

    # Load local last commit SHA
    $localCommitSha = if (Test-Path $lastCommitShaFile) { (Get-Content $lastCommitShaFile -Raw).Trim() } else { "" }

    if ($remoteCommitSha -eq $localCommitSha) {
        Write-Host "No changes detected. Skipping sync." -ForegroundColor Green
        return
    }

    if ($localCommitSha -eq "") {
        Write-Host "First run, setting commit SHA. Skipping sync." -ForegroundColor Green
        $remoteCommitSha | Set-Content $lastCommitShaFile
        return
    }

    Write-Host "Changes detected. Proceeding with sync..." -ForegroundColor Yellow

    # Compare commits
    $compareUrl = "https://api.github.com/repos/$owner/$repo/compare/$($localCommitSha)...$($remoteCommitSha)"
    Write-Host "Compare URL: $compareUrl" -ForegroundColor Cyan
    $compareResponse = Invoke-GitHubApi -url $compareUrl -headers $apiHeaders -pat $pat -patPath $patPath
    $compareData = ConvertFrom-Json $compareResponse.Content
    Write-Host "Files to sync: $($compareData.files.Count)" -ForegroundColor Green

    $updatedFiles = 0
    $scriptUpdated = $false
    foreach ($file in $compareData.files) {
        $path = $file.filename
        $status = $file.status
        $fullPath = "$PSScriptRoot\$path"

        if ($status -eq "removed") {
            if (Test-Path $fullPath) {
                Remove-Item $fullPath -Force
                Write-Host "Deleted: $path" -ForegroundColor Yellow
                Write-Log "Deleted: $path"
                $updatedFiles++
            }
        } elseif ($status -eq "added" -or $status -eq "modified") {
            $downloadUrl = $file.raw_url
            $downloadHeaders = if ($pat) { @{ Authorization = "Bearer $pat" } } else { @{} }
            try {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $fullPath -Headers $downloadHeaders -UseBasicParsing
                Write-Host "Downloaded/Updated: $path" -ForegroundColor Green
                Write-Log "Downloaded/Updated: $path"
                $updatedFiles++
                if ($path -eq "VCATechManager.ps1") {
                    $scriptUpdated = $true
                }
            } catch {
                Write-Host "Failed to download $path : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # Update last commit SHA
    $remoteCommitSha | Set-Content $lastCommitShaFile
    Write-Host "Repo sync complete. Updated $updatedFiles files." -ForegroundColor Green
    Write-Log "Repo sync complete"
    if ($scriptUpdated) {
        Write-Host "VCATechManager.ps1 has been updated. Relaunching in 3 seconds..." -ForegroundColor Cyan
        Write-Log "Relaunching script due to update"
        Start-Sleep -Seconds 3
        Start-Process -FilePath "$PSScriptRoot\VCATechManager.cmd" -WorkingDirectory $PSScriptRoot
        exit
    }
}

# Set global script root for use in functions
$global:ScriptRoot = $PSScriptRoot

try {
    # Import the custom module for shared functions (e.g., Get-UserSessionsParallel)
    Import-Module -Name "$PSScriptRoot\Private\VCATechManagerFunctions.psm1" -ErrorAction Stop
    Write-Log "Imported custom module: VCATechManagerFunctions"

    # New: Auto-update logic in the version check section
    try {
        $remoteVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/marcky168/VCATechManager/main/version.txt" -UseBasicParsing -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Content
        if ($remoteVersion -gt $version) {
            Write-Host "New version available: $remoteVersion. Update recommended." -ForegroundColor Yellow
            Write-Log "New version detected: $remoteVersion"
            $updateChoice = Read-Host "Download and update to version $remoteVersion? (y/n)"
            if ($updateChoice.ToLower() -eq 'y') {
                try {
                    $scriptUrl = "https://raw.githubusercontent.com/marcky168/VCATechManager/main/VCATechManager.ps1"
                    $newScriptPath = "$PSScriptRoot\VCATechManager_new.ps1"
                    Invoke-WebRequest -Uri $scriptUrl -OutFile $newScriptPath -UseBasicParsing
                    Move-Item -Path $newScriptPath -Destination $PSScriptRoot\VCATechManager.ps1 -Force
                    Write-Host "Updated to version $remoteVersion. Restart the script to use the new version." -ForegroundColor Green
                    Write-Log "Updated script to version $remoteVersion"
                } catch {
                    Write-Host "Update failed: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Update failed: $($_.Exception.Message)"
                }
            }
        }

        # Check for repo updates using version check
        Write-Host "Checking for repo updates..." -ForegroundColor Yellow
        $remoteVersionUrl = "https://raw.githubusercontent.com/marcky168/VCATechManager/main/Private/Version.txt"
        try {
            $remoteVersion = Invoke-WebRequest -Uri $remoteVersionUrl -UseBasicParsing | Select-Object -ExpandProperty Content
            $remoteVersion = $remoteVersion.Trim()
        } catch {
            Write-Host "Failed to check remote version: $($_.Exception.Message)" -ForegroundColor Yellow
            $remoteVersion = $version
        }
        if ($remoteVersion -gt $version) {
            Write-Host "New version available: $remoteVersion (local: $version). Running sync..." -ForegroundColor Green
            try {
                Sync-Repo
            } catch {
                Write-Host "Repo update failed: $($_.Exception.Message)" -ForegroundColor Red
                Write-Log "Repo update failed: $($_.Exception.Message)"
            }
        } else {
            Write-Host "No new version available (remote: $remoteVersion, local: $version). Skipping sync." -ForegroundColor Green
        }
    } catch {
        Write-Host "Failed to check for updates: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log "Failed to check for updates: $($_.Exception.Message)"
    }

    # Import ActiveDirectory module with check
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Host "ActiveDirectory module loaded successfully." -ForegroundColor Cyan  # Debug
    } catch {
        Write-Host "ActiveDirectory module failed to load. Install RSAT." -ForegroundColor Red
        Write-Log "ActiveDirectory import error: $($_.Exception.Message)"
    }
    # Check and install PoshRSJob if missing
    $poshLibPath = "$PSScriptRoot\Private\lib"
    $poshVersionPath = "$poshLibPath\PoshRSJob\1.7.4.4"
    if (-not (Test-Path $poshVersionPath)) {
        Write-Host "PoshRSJob module not found at $poshVersionPath. Downloading and installing locally..." -ForegroundColor Yellow
        Write-Log "PoshRSJob not found. Starting local installation."
        try {
            # Ensure lib folder exists
            if (-not (Test-Path $poshLibPath)) {
                New-Item -Path $poshLibPath -ItemType Directory -Force | Out-Null
            }
            
            # Download module from PowerShell Gallery to lib folder
            Save-Module -Name PoshRSJob -RequiredVersion 1.7.4.4 -Path $poshLibPath -Force -ErrorAction Stop
            
            Write-Host "PoshRSJob (v1.7.4.4) installed locally to $poshVersionPath." -ForegroundColor Green
            Write-Log "PoshRSJob installed successfully to $poshVersionPath."
        } catch {
            Write-Host "Failed to install PoshRSJob: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "PoshRSJob installation error: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
            # Optionally exit or continue without the module
        }
    }
    # Import required modules with try-catch
    try {
        Import-Module -Name "$PSScriptRoot\Private\lib\PoshRSJob\1.7.4.4\PoshRSJob.psm1" -ErrorAction Stop
        Write-Log "Imported module: PoshRSJob"
        Import-Module -Name "$PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction Stop
        Write-Log "Imported module: PSTerminalServices"
        Import-Module -Name "$PSScriptRoot\Private\lib\ImportExcel" -ErrorAction Stop  # Added for HOSPITALMASTER loading
        Write-Log "Imported module: ImportExcel"
        # Suppress PnP PowerShell update check
        $env:PNPPOWERSHELL_UPDATECHECK = 'Off'
        Import-Module -Name "$PSScriptRoot\Private\lib\PnP.PowerShell" -ErrorAction Stop  # Added for SharePoint access
        Write-Log "Imported module: PnP.PowerShell"
        Import-Module -Name "$PSScriptRoot\Private\lib\Posh-SSH" -ErrorAction Stop  # Added for SSH functionality
        Write-Log "Imported module: Posh-SSH"
        Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Sdk" -ErrorAction Stop  # Added for VMware functionality
        Write-Log "Imported module: VMware.VimAutomation.Sdk"
        Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Common" -ErrorAction Stop  # Added for VMware functionality
        Write-Log "Imported module: VMware.VimAutomation.Common"
        Import-Module -Name "$PSScriptRoot\Private\lib\VMware.Vim" -ErrorAction Stop  # Added for VMware functionality
        Write-Log "Imported module: VMware.Vim"
        Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Cis.Core" -ErrorAction Stop  # Added for VMware functionality
        Write-Log "Imported module: VMware.VimAutomation.Cis.Core"
        Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Core" -ErrorAction Stop  # Added for VMware functionality
        Write-Log "Imported module: VMware.VimAutomation.Core"
        Import-Module -Name "$PSScriptRoot\Private\lib\Autoload" -Verbose
        Write-Log "Imported module: Autoload"
        Write-Host "Required modules loaded successfully." -ForegroundColor Cyan  # Debug
        Write-Log "Required modules loaded successfully."
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
        $hospitalMasterPath = "$PSScriptRoot\Data\HOSPITALMASTER.xlsx"
        $hospitalMasterUrl = $config.InternalDomains.SharePointBaseUrl + $config.InternalDomains.HospitalMasterPath

        if (Test-Path $hospitalMasterPath) {
            try {
                $HospitalMaster = Import-Excel -Path $hospitalMasterPath -WorksheetName Misc
                Write-Log "Hospital master loaded successfully from $hospitalMasterPath"
            } catch {
                Write-Host "Failed to load hospital master: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Log "Hospital master load error: $($_.Exception.Message)"
            }
        } else {
            # Attempt automatic download for VCA employees with valid credentials
            if ($ADCredential -and (Test-ADCredentials -Credential $ADCredential)) {
                Write-Host "Hospital master file not found locally. Attempting secure download..." -ForegroundColor Cyan
                try {
                    # Ensure Data directory exists
                    if (-not (Test-Path "$PSScriptRoot\Data")) {
                        New-Item -ItemType Directory -Path "$PSScriptRoot\Data" -Force | Out-Null
                    }

                    # Download the file using credentials
                    Invoke-WebRequest -Uri $hospitalMasterUrl -OutFile $hospitalMasterPath -Credential $ADCredential -UseBasicParsing
                    Write-Host "Hospital master downloaded successfully from secure VCA server." -ForegroundColor Green
                    Write-Log "Hospital master downloaded from $hospitalMasterUrl"

                    # Load the downloaded file
                    $HospitalMaster = Import-Excel -Path $hospitalMasterPath -WorksheetName Misc
                    Write-Log "Hospital master loaded successfully after download"
                } catch {
                    Write-Host "Failed to download hospital master: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "Please download the HOSPITALMASTER.xlsx file manually from the secure internal file share and place it in the 'Data' folder." -ForegroundColor Cyan
                    Write-Log "Hospital master download failed: $($_.Exception.Message)"
                }
            } else {
                Write-Host "Hospital master file not found at $hospitalMasterPath." -ForegroundColor Yellow
                Write-Host "Please download the HOSPITALMASTER.xlsx file from the secure internal file share and place it in the 'Data' folder." -ForegroundColor Cyan
                Write-Host "Note: Automatic download requires valid VCA domain credentials." -ForegroundColor Gray
                Write-Log "Hospital master file not found and no valid credentials for automatic download"
            }
            # Don't exit - allow script to continue without hospital data
        }
    }

    # Function to get servers for AU with enhanced validation and error handling
    function Get-VCAServers {
        param([string]$AU)
        # Parameter validation: AU must be numeric and 3-6 digits
        if ($AU -notmatch '^\d{3,6}$') {
            throw "Invalid AU number. Please enter a 3 to 6 digit number."
        }
        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''
        try {
            # Use splatting for Get-ADComputer
            $adComputerParams = @{
                Filter      = "Name -like '$SiteAU-ns*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'"
                Server      = $config.InternalDomains.PrimaryDomain
                Credential  = $ADCredential
                ErrorAction = 'Stop'
            }
            $adServers = Get-ADComputer @adComputerParams | Select-Object -ExpandProperty Name | Sort-Object Name
            if (-not $adServers -or $adServers.Count -eq 0) {
                throw "No -NS servers found for AU $AU."
            }
            return $adServers
        } catch {
            if ($_.Exception.Message -like "*credentials*") {
                Write-Host "Warning: AD credentials invalid. Some features may not work. Update via menu option 11." -ForegroundColor Yellow
                Write-Log "AD credentials invalid in Get-VCAServers: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
                return @()
            } else {
                Write-Log "Error fetching servers for AU $AU : $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
                throw
            }
        }
    }

    # New: Function for Abaxis MAC Address Search (optimized: uses helpers, caches IP, splatting)
    function Invoke-AbaxisMacSearch {
        param([string]$AU)
        Write-Log "Starting Abaxis MAC Address Search for AU $AU"

        # Cache for IP resolutions (optimization: avoid repeated DNS calls)
        $ipCache = @{}

        $dhcpServer = $config.NetworkSettings.PrimaryDHCPServer
        $hostname = Convert-VcaAu -AU $AU -Suffix '-gw'

        # Resolve gateway IP using helper function
        $ipAddresses = Resolve-HostIP -Hostname $hostname -IpCache $ipCache
        if (-not $ipAddresses -or $ipAddresses.Length -eq 0) {
            Write-Host "Error: No IP addresses found for hostname '$hostname'." -ForegroundColor Red
            Write-Log "No IP found for $hostname"
            return
        } elseif ($ipAddresses.Length -gt 1) {
            Write-Host "Warning: Multiple IP addresses found for '$hostname'. Using the first one: $($ipAddresses[0].ToString())" -ForegroundColor Yellow
        }
        $ip = $ipAddresses[0].ToString()
        $scopeId = $ip -replace '\.\d+$', '.0'

        # Retrieve DHCP leases using helper function
        Write-Progress -Activity "Retrieving DHCP leases" -Status "Connecting to $dhcpServer..." -PercentComplete 50
        $leases = Get-DHCPLeases -DHCPServer $dhcpServer -ScopeId $scopeId
        if (-not $leases) {
            Write-Host "No leases found for scope '$scopeId'."
        }

        # Process each group and find matching leases
        $groupResults = @()
        $macPrefixes = @{
            "VS2"   = @("00-07-32", "00-30-64")
            "HM5"   = @("00-1B-EB")
            "VSPro" = @("00-03-1D")
            "Fuse"  = @("00-90-FB", "00-50-56", "00-0C-29")
        }

        foreach ($group in $macPrefixes.Keys) {
            $prefixes = $macPrefixes[$group]
            $normalizedPrefixes = $prefixes | ForEach-Object { Normalize-MacAddress $_ }

            $matchingLeases = $leases | Where-Object {
                $normalizedClientId = Normalize-MacAddress $_.ClientId
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

        # Retrieve and display DHCP reservations for the scope using helper function
        Write-Host "`nDHCP Reservations for scope $scopeId" -ForegroundColor Green
        $reservations = Get-DHCPReservations -DHCPServer $dhcpServer -ScopeId $scopeId
        if ($reservations) {
            $reservations | Sort-Object IPAddress | Format-Table -Property IPAddress, ClientId, Name, Description, @{Name="DeviceGroup"; Expression={Get-DeviceGroup $_.ClientId}}
            $groupResults += $reservations
        } else {
            Write-Host "No reservations found for scope '$scopeId'."
        }

        # Retrieve ARP data from servers using helper function
        Write-Host "`nRetrieving ARP data from servers..." -ForegroundColor Green
        $servers = Get-CachedServers -AU $AU
        $arpResults = @()
        if ($servers) {
            $arpResults = Get-ARPTable -ComputerNames $servers -Credential $ADCredential
        }        # Process ARP results and add matching devices to groupResults (excluding Fuse)
        $uniqueArp = $arpResults | Sort-Object IPAddress -Unique
        foreach ($entry in $uniqueArp) {
            $group = Get-DeviceGroup $entry.LinkLayerAddress
            if ($group -ne "Other" -and $group -ne "Fuse") {
                Write-Host "ARP Device: $($entry.IPAddress) - $group" -ForegroundColor Cyan
                $arpObject = [PSCustomObject]@{
                    IPAddress   = $entry.IPAddress
                    ClientId    = $entry.LinkLayerAddress
                    DeviceGroup = $group
                    Source      = "ARP"
                }
                $groupResults += $arpObject
            }
        }

        # Ping other leased devices (excluding Fuse)
        $otherDevices = $groupResults | Where-Object { $_.ClientId -notmatch '^00-90-FB|^00-50-56|^00-0C-29' }
        $runPingTest = Read-Host "Run ping test on other leased devices? (y/n)"
        if ($runPingTest.ToLower() -eq 'y') {
            Write-Host "Testing connectivity to $($otherDevices.Count) devices (parallel processing)..." -ForegroundColor Green
            $pingJobs = @()
            foreach ($device in $otherDevices) {
                $job = Start-RSJob -ScriptBlock {
                    param($device)
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
                    [PSCustomObject]@{
                        DeviceName = $deviceName
                        IP = $ip
                        PingResult = $pingResult
                    }
                } -ArgumentList $device
                $pingJobs += $job
            }

            # Wait for all ping jobs and display results
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
            Write-Host "`nFuse Device IP $fuseType from nslookup on $fuseHostname : " -ForegroundColor Green -NoNewline
            Write-Host "$fuseIp" -ForegroundColor Yellow
            $pingResult = Test-Connection -ComputerName $fuseIp -Count 4 -ErrorAction SilentlyContinue
            if ($pingResult) {
                $pingResult | Format-Table -Property Address, ResponseTime, StatusCode
                Write-Host "Fuse device is responsive." -ForegroundColor Green
            } else {
                Write-Host "Fuse device did not respond to ping." -ForegroundColor Red
            }
            # Prompt to open Fuse webpage regardless of ping status
            $openFuse = Read-Host "Do you want to open the Fuse webpage? (y/n)"
            if ($openFuse.ToLower() -eq 'y') {
                $fuseUrl = "https://${fuseHostname}:8443"
                Start-Process "msedge" -ArgumentList $fuseUrl
                Write-Host "Opening Fuse webpage: $fuseUrl" -ForegroundColor Green
            }
            # Offer vSphere reboot if not responding and virtual
            if (-not $pingResult -and $fuseIp -like "10.242*") {
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
            # Use splatting for Start-RSJob
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server)
                    try {
                        # Use splatting for Invoke-Command
                        $invokeParams = @{
                            ComputerName = $server
                            SessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                            ScriptBlock = {
                                $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
                                $serverTime = $using:server + '  ' + $time

                                try {
                                    $allErrors = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2} -MaxEvents 200 -ErrorAction Stop
                                } catch {
                                    Write-Debug "Failed to query Woofware errors on $env:COMPUTERNAME: $($_.Exception.Message)"
                                    $allErrors = @()
                                }

                                $allErrors
                            }
                        }
                        Invoke-Command @invokeParams
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

        # Process results to match 70w columns
        $processedResults = $results | ForEach-Object {
            $msg = $_.Message

            $threadIdentity  = if ($msg -match 'ThreadIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $windowsIdentity = if ($msg -match 'WindowsIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $machineName     = if ($msg -match 'MachineName\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $exceptionType   = if ($msg -match 'Type\s*:\s*([^,]+)') { $matches[1].Trim() } else { 'N/A' }
            $MsgError        = if ($msg -match '(?s)Message\s*:\s*(AUID\s*=\s*.*?)(?:\r?\n\S|\Z)') { $matches[1].Trim() } else { 'N/A' }

            [PSCustomObject]@{
                PSComputerName     = $_.PSComputerName
                TimeCreated        = $_.TimeCreated
                EventID            = $_.Id
                RecordID           = $_.RecordId
                MachineName        = $machineName
                ThreadIdentity     = $threadIdentity
                WindowsIdentity    = $windowsIdentity
                ExceptionType      = $exceptionType
                MessagError        = $MsgError
                FullMessage        = $msg
            }
        }

        # Display all errors in one grid
        $selectedError = $processedResults | Out-GridView -Title "#2 Woofware Errors Check - v.$version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

        # Display selected error details
        if ($selectedError) {
            Write-Host "Selected Error Details:" -ForegroundColor Cyan
            $selectedError | Format-List MachineName, TimeCreated, EventID, RecordID, ThreadIdentity, WindowsIdentity, ExceptionType, FullMessage

            # Prompt to copy error message to clipboard
            $copyToClipboard = Read-Host "Copy error message to clipboard? (y/n)"
            if ($copyToClipboard.ToLower() -eq 'y') {
                $selectedError.FullMessage | Set-Clipboard
                Write-Host "Error message copied to clipboard." -ForegroundColor Green
            }
        }

        # Use helper for export
        Export-Results -Results $processedResults -BaseName "woofware_results" -AU $AU

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
                $errorDetails = "Selected Error Details:`n" +
                    "Server: $($selectedError.PSComputerName)`n" +
                    "Time Created: $($selectedError.TimeCreated)`n" +
                    "Event ID: $($selectedError.EventID)`n" +
                    "Record ID: $($selectedError.RecordID)`n" +
                    "Machine Name: $($selectedError.MachineName)`n" +
                    "Thread Identity: $($selectedError.ThreadIdentity)`n" +
                    "Windows Identity: $($selectedError.WindowsIdentity)`n" +
                    "Exception Type: $($selectedError.ExceptionType)`n" +
                    "Message Error: $($selectedError.MessagError)`n" +
                    "Full Message: $($selectedError.FullMessage)"
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

            # Ensure Outlook is running
            Start-OutlookIfNeeded

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
                        $signatureHtml = $signatureHtml -replace "cid:$cid", ('data:' + $mime + ';base64,' + $base64)
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
            $bodyHtml = [string]::Format('<p><strong style="color: red;">Team,</strong></p><p><strong>Details:</strong></p><p><span style="font-weight: bold; color: red;">Location:</span> <span style="color: #4169e1;">{0}</span></p><p><span style="font-weight: bold; color: red;">Site Contact Name:</span> <span style="color: #4169e1;">{1} ({2})</span></p><p><span style="font-weight: bold; color: red;">Issue:</span> <span style="color: #4169e1;">{3}</span></p><p><strong style="color: red;">Error Details:</strong></p><pre>{4}</pre><br><br>{5}', $location, $contact, $phone, $description, $errorDetails, $signatureHtml)

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
            $groupMembers = Get-ADGroupMember -Identity $adGroupName -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
            $adUsers = $groupMembers | Get-ADUser -Properties Name, SamAccountName -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop |
                       Select-Object Name, SamAccountName | Sort-Object Name
            if ($adUsers) {
                $selectedUser = $adUsers | Out-GridView -Title "Select user to filter Woofware errors for AU $AU" -OutputMode Single
                if ($selectedUser) {
                    $Username = $selectedUser.SamAccountName
                }
            } else {
                Write-Host "No users found in AD group $adGroupName." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Failed to query AD group '$adGroupName' for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "AD query error in Woofware-ErrorsCheckByUser: $($_.Exception.Message)"
        }

        # Prompt for username if not selected or no users
        if (-not $Username) {
            $Username = Read-Host "Enter username to search for Woofware errors"
            if (-not $Username) {
                Write-Host "No username provided. Cancelling." -ForegroundColor Yellow
                return
            }
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
            # Use splatting for Start-RSJob
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server, $Username)
                    try {
                        $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                        Invoke-Command -ComputerName $server -SessionOption $sessionOption -ScriptBlock {
                            $time = (Get-CimInstance win32_operatingsystem).LocalDateTime
                            $serverTime = $using:server + '  ' + $time

                            try {
                                $allErrors = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2} -MaxEvents 200 -ErrorAction Stop
                            } catch {
                                Write-Debug "Failed to query Woofware errors on $env:COMPUTERNAME: $($_.Exception.Message)"
                                $allErrors = @()
                            }

                            $allErrors
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
        Write-Progress -Activity "Querying Woofware errors by user" -Completed

        # Now filter by username in the main function
        $filteredErrors = $results | Where-Object { $_.Message -ilike "*VCAANTECH\$Username*" }

        # Process filtered errors to match 70w columns
        $processedResults = $filteredErrors | ForEach-Object {
            $msg = $_.Message

            $threadIdentity  = if ($msg -match 'ThreadIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $windowsIdentity = if ($msg -match 'WindowsIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $machineName     = if ($msg -match 'MachineName\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $exceptionType   = if ($msg -match 'Type\s*:\s*([^,]+)') { $matches[1].Trim() } else { 'N/A' }
            $MsgError        = if ($msg -match '(?s)Message\s*:\s*(AUID\s*=\s*.*?)(?:\r?\n\S|\Z)') { $matches[1].Trim() } else { 'N/A' }

            [PSCustomObject]@{
                PSComputerName     = $_.PSComputerName
                TimeCreated        = $_.TimeCreated
                EventID            = $_.Id
                RecordID           = $_.RecordId
                MachineName        = $machineName
                ThreadIdentity     = $threadIdentity
                WindowsIdentity    = $windowsIdentity
                ExceptionType      = $exceptionType
                MessagError        = $MsgError
                FullMessage        = $msg
            }
        }

        # Display all errors in a single grid for selection
        if ($processedResults) {
            $selectedError = $processedResults | Out-GridView -Title "#2b Woofware Errors Check by User - v.$version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single
            if ($selectedError) {
                Write-Host "Selected Error Details:" -ForegroundColor Cyan
                $selectedError | Format-List MachineName, TimeCreated, EventID, RecordID, ThreadIdentity, WindowsIdentity, ExceptionType, FullMessage

                # Prompt to copy error message to clipboard
                $copyToClipboard = Read-Host "Copy error message to clipboard? (y/n)"
                if ($copyToClipboard.ToLower() -eq 'y') {
                    $selectedError.FullMessage | Set-Clipboard
                    Write-Host "Error message copied to clipboard." -ForegroundColor Green
                }
            } else {
                Write-Host "No error selected." -ForegroundColor Yellow
            }
        } else {
            Write-Host "No Woofware errors found for user $Username in AU $AU." -ForegroundColor Yellow
            return
        }

        # Export using helper
        Export-Results -Results $processedResults -BaseName "woofware_user_results" -AU $AU

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
                $errorDetails = "Selected Error Details:`n" +
                    "Server: $($selectedError.PSComputerName)`n" +
                    "Time Created: $($selectedError.TimeCreated)`n" +
                    "Event ID: $($selectedError.EventID)`n" +
                    "Record ID: $($selectedError.RecordID)`n" +
                    "Machine Name: $($selectedError.MachineName)`n" +
                    "Thread Identity: $($selectedError.ThreadIdentity)`n" +
                    "Windows Identity: $($selectedError.WindowsIdentity)`n" +
                    "Exception Type: $($selectedError.ExceptionType)`n" +
                    "Message Error: $($selectedError.MessagError)`n" +
                    "Full Message: $($selectedError.FullMessage)"
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

            # Ensure Outlook is running
            Start-OutlookIfNeeded

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
                        $signatureHtml = $signatureHtml -replace "cid:$cid", ('data:' + $mime + ';base64,' + $base64)
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
            $bodyHtml = [string]::Format('<p><strong style="color: #CD5C5C;">Team,</strong></p><p><strong>Details:</strong></p><p><span style="font-weight: bold; color: #CD5C5C;">Location:</span> <span style="color: #4169e1;">{0}</span></p><p><span style="font-weight: bold; color: #CD5C5C;">Site Contact Name:</span> <span style="color: #4169e1;">{1} ({2})</span></p><p><span style="font-weight: bold; color: #CD5C5C;">Issue:</span> <span style="color: #4169e1;">{3}</span></p><p><strong style="color: #CD5C5C;">Error Details:</strong></p><pre>{4}</pre><br><br>{5}', $location, $contact, $phone, $description, $errorDetails, $signatureHtml)

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

    # New: Function for Application Hang Errors Check
    function ApplicationHang-ErrorsCheck {
        param([string]$AU)
        Write-Log "Starting Application Hang Errors Check for AU $AU"

        try {
            $servers = Get-CachedServers -AU $AU
        } catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Write-Log "Error in Application Hang check: $($_.Exception.Message)"
            return
        }

        $jobs = @()
        $totalServers = $servers.Count
        $i = 0
        foreach ($server in $servers) {
            $i++
            Write-Progress -Activity "Querying Application Hang errors" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)
            # Use splatting for Start-RSJob
            $jobParams = @{
                Name         = $server
                ScriptBlock  = {
                    param($server)
                    try {
                        # Use splatting for Invoke-Command
                        $invokeParams = @{
                            ComputerName = $server
                            SessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                            ScriptBlock = {
                                try {
                                    $allErrors = Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Application Hang'; Id=1002} -MaxEvents 200 -ErrorAction Stop
                                } catch {
                                    Write-Debug "Failed to query Application Hang errors on $env:COMPUTERNAME: $($_.Exception.Message)"
                                    $allErrors = @()
                                }

                                $allErrors
                            }
                        }
                        Invoke-Command @invokeParams
                    } catch {
                        Write-Debug "Error querying sessions on $server : $($_.Exception.Message)"
                        @()
                    }
                }
                ArgumentList = $server
            }
            $jobs += Start-RSJob @jobParams
        }

        $results = $jobs | Wait-RSJob | ForEach-Object { Receive-RSJob -Job $_; Remove-RSJob -Job $_ } | Where-Object { $_ }
        Write-Progress -Activity "Querying Application Hang errors" -Completed

        # Process results to match 70w columns
        $processedResults = $results | ForEach-Object {
            $msg = $_.Message

            $threadIdentity  = if ($msg -match 'ThreadIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $windowsIdentity = if ($msg -match 'WindowsIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $machineName     = if ($msg -match 'MachineName\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
            $exceptionType   = if ($msg -match 'Type\s*:\s*([^,]+)') { $matches[1].Trim() } else { 'N/A' }
            $MsgError        = if ($msg -match '(?s)Message\s*:\s*(AUID\s*=\s*.*?)(?:\r?\n\S|\Z)') { $matches[1].Trim() } else { 'N/A' }

            [PSCustomObject]@{
                PSComputerName     = $_.PSComputerName
                TimeCreated        = $_.TimeCreated
                EventID            = $_.Id
                RecordID           = $_.RecordId
                MachineName        = $machineName
                ThreadIdentity     = $threadIdentity
                WindowsIdentity    = $windowsIdentity
                ExceptionType      = $exceptionType
                MessagError        = $MsgError
                FullMessage        = $msg
            }
        }

        # Display all errors in one grid
        $selectedError = $processedResults | Out-GridView -Title "#2c Application Hang Errors Check - v.$version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

        # Display selected error details
        if ($selectedError) {
            Write-Host "Selected Error Details:" -ForegroundColor Cyan
            $selectedError | Format-List MachineName, TimeCreated, EventID, RecordID, ThreadIdentity, WindowsIdentity, ExceptionType, FullMessage

            # Prompt to copy error message to clipboard
            $copyToClipboard = Read-Host "Copy error message to clipboard? (y/n)"
            if ($copyToClipboard.ToLower() -eq 'y') {
                $selectedError.FullMessage | Set-Clipboard
                Write-Host "Error message copied to clipboard." -ForegroundColor Green
            }
        }

        # Use helper for export
        Export-Results -Results $processedResults -BaseName "application_hang_results" -AU $AU

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

            $subject = "AU$($AU.PadLeft(4, '0')) Application Hang Error"

            # Build error details string
            if ($selectedError) {
                $errorDetails = "Selected Error Details:`n" +
                    "Server: $($selectedError.PSComputerName)`n" +
                    "Time Created: $($selectedError.TimeCreated)`n" +
                    "Event ID: $($selectedError.EventID)`n" +
                    "Record ID: $($selectedError.RecordID)`n" +
                    "Machine Name: $($selectedError.MachineName)`n" +
                    "Thread Identity: $($selectedError.ThreadIdentity)`n" +
                    "Windows Identity: $($selectedError.WindowsIdentity)`n" +
                    "Exception Type: $($selectedError.ExceptionType)`n" +
                    "Message Error: $($selectedError.MessagError)`n" +
                    "Full Message: $($selectedError.FullMessage)"
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

            # Ensure Outlook is running
            Start-OutlookIfNeeded

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
                        $signatureHtml = $signatureHtml -replace "cid:$cid", ('data:' + $mime + ';base64,' + $base64)
                    } catch {
                        # Ignore attachment processing errors
                    }
                }

                # Clean up dummy email
                $dummyMail.Close(1)  # 1 = olDiscard
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($dummyMail) | Out-Null
                $dummyMail = $null
            } catch {
                Write-Host "Failed to capture Outlook signature: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Log "Signature capture error: $($_.Exception.Message)"
            }

            # Build HTML body
            $bodyHtml = [string]::Format('<html><body><p>Dear Dev Team,</p><p>{0}</p><p>{1}</p><p>Best regards,<br>{2}</p>{3}</body></html>', $description, $errorDetails, $(whoami), $signatureHtml)

            try {
                # Create real email
                $mail = $outlook.CreateItem(0)

                # Set email properties
                $mail.To = $to
                if ($cc) { $mail.CC = $cc }
                $mail.Subject = $subject
                $mail.HTMLBody = $bodyHtml  # Set body with embedded signature images

                $mail.Display()  # Opens as draft for review/edit/send
                Write-Host "Email draft created in Outlook with embedded images." -ForegroundColor Green
                Write-Log "Created email draft for Application Hang errors AU $AU"

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
        $totalServers = 2  # Assuming 2 DHCP servers as per context
        $i = 0
        foreach ($Server in $config.NetworkSettings.DHCPServers) {
            $i++
            Write-Progress -Activity "Processing DHCP servers" -Status "Server $i of $totalServers : $Server" -PercentComplete (($i / $totalServers) * 100)
            try {
                $ExistingReservation = Get-DhcpServerv4Reservation -ComputerName $Server -IPaddress $ReservationIP -ErrorAction Stop
                if ($ExistingReservation) {
                    $Confirm = Read-Host "A DHCP reservation with IP address $ReservationIP and scope $ScopeId already exists on server $Server. Do you want to delete it? (y/n)"
                    if ($Confirm.ToLower() -eq "y") {
                        Remove-DhcpServerv4Reservation -ComputerName $Server -IPAddress $ReservationIP -ErrorAction Stop
                        $results += "Deleted DHCP reservation for IP address $ReservationIP and scope $ScopeId on server $Server"
                        Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for AU $AU Fuse"
                        $results += "Added DHCP reservation for IP address $ReservationIP to scope $ScopeId on server $Server"
                    }
                } else {
                    Add-DhcpServerv4Reservation -ComputerName $Server -ScopeId $ScopeId -IPAddress $ReservationIP -ClientId $MACAddress -Description "Reservation for AU $AU Fuse"
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

        # Use helper
        Export-Results -Results $results -BaseName "dhcp_results" -AU $AU
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
                    # Launch VNC using helper function
                    $userIP = $SelectedUser.IPAddress
                    Start-VNCViewer -IPAddress $userIP -Username $SelectedUser.UserName -Computer $SelectedUser.Computer
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
                $groupMembers = Get-ADGroupMember -Identity $adGroupName -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
                $adUsers = $groupMembers | Get-ADUser -Properties Name, SamAccountName, EmailAddress, Title -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop | 
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
                    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential).MaxPasswordAge
                    $adUser = Get-ADUser -Identity $Username -Properties Name, Title, OfficePhone, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop
                    Write-Host "`nAD Properties for $Username :" -ForegroundColor Cyan
                    $adUser | Select-Object Name, Title, @{n='OfficePhone'; e={$_.OfficePhone}}, Office, Department, EmailAddress, StreetAddress, City, State, PostalCode, SID, Created, extensionAttribute3, PasswordLastSet, @{n='PasswordExpires'; e={ if ($_.PasswordLastSet) { $_.PasswordLastSet + $MaxPasswordAge } else { 'Never Set' } }} | Format-List
                } catch {
                    Write-Host "User '$Username' not found in AD. Proceed anyway? (y/n)" -ForegroundColor Yellow
                    if ((Read-Host).ToLower() -ne 'y') { return }
                }
            }

            try {
                $servers = Get-CachedServers -AU $AU -ErrorAction Stop
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
                        # Launch VNC using helper function
                        $userIP = $selectedSession.ClientIP
                        Start-VNCViewer -IPAddress $userIP -Username $selectedSession.UserName -Computer $selectedSession.Server
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

        $HospitalMasterUrl = $config.InternalDomains.SharePointBaseUrl + $config.InternalDomains.HospitalMasterPath
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
            $SharePointUrl = $config.InternalDomains.SharePointBaseUrl + '/sites/WOOFconnect/regions'
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
            $groupMembers = Get-ADGroupMember -Identity $groupName -Server $config.InternalDomains.PrimaryDomain -Credential $Credential -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' }
            $users = $groupMembers | Get-ADUser -Properties Name, SamAccountName, EmailAddress, LockedOut, PasswordExpired, LastLogonDate -Server $config.InternalDomains.PrimaryDomain -Credential $Credential -ErrorAction Stop

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
            $AngryIPPath = "$PSScriptRoot\Private\bin\ipscan-win64-3.9.2.exe"

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

# Main script logic with menu, enhanced AU validation and try-catch
    $exitScript = $false

    while (-not $exitScript) {
        # Reset AU to ensure prompt is shown
        $AU = $null
        Clear-Host

        # Display tool name and version at top with spacing
        Write-Host "`n`n  VCATechManager v$version`n" -ForegroundColor Magenta

        Write-Host "Enter the AU number (or 'exit' to quit): " -NoNewline -ForegroundColor Cyan
        $AU = (Read-Host).Trim()
        
        if ($AU -eq 'exit') {
            $exitScript = $true
            continue
        }

        # Enhanced AU validation
        if ($AU -notmatch '^\d{3,6}$') {
            Write-Host "Invalid AU number. Please enter a 3 to 6 digit number." -ForegroundColor Red
            Start-Sleep -Seconds 2
            continue
        }

        try {
            if ($validAUs[$AU]) {
                $servers = $validAUs[$AU]
            } else {
                $servers = Get-VCAServers -AU $AU -ErrorAction Stop
                $validAUs[$AU] = $servers
            }
            # Check for empty servers
            if (-not $servers -or $servers.Count -eq 0) {
                Write-Host "No servers found for AU $AU. Try another AU." -ForegroundColor Red
                continue
            }
        } catch {
            Write-Host "AU $AU invalid or no servers found. Try another? Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "AU validation error for $AU : $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
            continue
        }

        # Validate AD group for AU with try-catch
        $adGroupName = 'H' + $AU.PadLeft(4, '0')
        Write-Log "Debug: Validating group name '$adGroupName' for AU $AU"
        try {
            $group = Get-ADObject -Filter "objectClass -eq 'group' -and name -eq '$adGroupName'" -Server $config.InternalDomains.PrimaryDomain -Credential $ADCredential -ErrorAction Stop
            if (-not $group) { throw "Group '$adGroupName' not found." }
        } catch {
            Write-Host "Failed to query AD group '$adGroupName' for AU $AU. Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Failed to query AD group '$adGroupName' for AU $AU : $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
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
        $host.UI.RawUI.WindowTitle = "VCATechManager v$version - [$AU] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available'}) - $scriptPath"

        # Display the menu once after entering AU
        Write-Host "`n--- Main Menu for AU $AU (v$version) ---" -ForegroundColor Green

        # Menu Improvements: Use hashtable for menu options
        $menuOptions = @{
            "0" = "Change AU"
            "1" = "Abaxis MAC Address Search"
            "2" = "Woofware Errors Check"
            "2b" = "Woofware Errors Check by User"
            "2c" = "Application Hang Errors Check"
            "3" = "Add DHCP Reservation"
            "4" = "GPUpdate /Force on Selected Server"
            "5" = "List AD Users and Check Logon"
            "6" = "Kill Sparky Shell for Logged-in User"
            "7" = "Exit"
            "8" = "Help"
            "9" = "Toggle Verbose Logging (Current): $(if ($verboseLogging) {'On'} else {'Off'})"
            "10" = "Robo Update"
            "11" = "Update Admin Credentials"
            "12" = "Device Connectivity Test"
            "13" = "Launch ServiceNow for AU Tickets"
            "14" = "AD User Management"
            "14u" = "Update Hospital Master"
            "15" = "Launch vSphere for Fuse VM"
            "16" = "Run Angry IP Scanner on DHCP Scope"
            "19" = "Launch Remote Desktop to Selected Servers"
            "999" = "Restart Script"
            "81" = "Launch WOOFware Reports Website"
            "82" = "Launch Fuse Website"
            "83" = "win: Restart Sparky Services"
        }

        # Sort keys numerically, handling non-numeric suffixes (e.g., "2b", "14u")
        foreach ($key in ($menuOptions.Keys | Sort-Object @{Expression={ 
            $match = [regex]::Match($_, '^(\d+)'); 
            if ($match.Success) { [int]$match.Groups[1].Value } else { 999 }
        }}, @{Expression={$_}})) {
            Write-Host "$key. $($menuOptions[$key])" -ForegroundColor Cyan
        }

        $menuActive = $true
        do {
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
                        "America/Phoenix" = "US Mountain Standard Time"
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
                    Write-Host "Returning to AU prompt..." -ForegroundColor Green
                    $menuActive = $false
                    # Reset window title to base
                    $host.UI.RawUI.WindowTitle = "VCATechManager v$version - $scriptPath"
                }
                "1" {
                    Invoke-MenuOption1 -AU $AU -ADCredential $ADCredential -credPathAD $credPathAD
                }
                "2" {
                    # Import PoshRSJob for parallel processing
                    try {
                        Import-Module -Name "$PSScriptRoot\Private\lib\PoshRSJob\1.7.4.4\PoshRSJob.psm1" -ErrorAction Stop
                    } catch {
                        Write-Host "PoshRSJob module required for Woofware check." -ForegroundColor Red
                        continue
                    }
                    Woofware-ErrorsCheck -AU $AU
                }
                "2b" {
                    Woofware-ErrorsCheckByUser -AU $AU
                }
                "2c" {
                    ApplicationHang-ErrorsCheck -AU $AU
                }
                "3" {
                    Add-DHCPReservation -AU $AU
                }
                "4" {
                    Invoke-GPUpdateForce -AU $AU
                }
                "5" {
                    # Check AD credentials before proceeding
                    if (-not $ADCredential -or -not (Test-ADCredentials -Credential $ADCredential)) {
                        Write-Host "AD credentials invalid or missing. Prompting for new ones..." -ForegroundColor Yellow
                        try {
                            $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)" -ErrorAction Stop
                            if ($ADCredential) {
                                $ADCredential | Export-Clixml -Path $credPathAD -Force -ErrorAction Stop
                                Write-Host "AD credentials saved." -ForegroundColor Green
                                Write-Log "AD credentials updated via option 5."
                            } else {
                                Write-Host "No credentials provided. Skipping List AD Users and Check Logon." -ForegroundColor Yellow
                                return
                            }
                        } catch {
                            Write-Host "Error updating credentials: $($_.Exception.Message)" -ForegroundColor Red
                            Write-Log "Error updating credentials in option 5: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
                            return
                        }
                    }
                    try {
                        User-LogonCheck -AU $AU -ErrorAction Stop
                    } catch {
                        Write-Host "Error in option 5: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Log "Error in option 5: $($_.Exception.Message) | StackTrace: $($_.Exception.StackTrace)"
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
                    Write-Host "19. Launch Remote Desktop to Selected Servers: Launches RDP to selected servers." -ForegroundColor White
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
                        $ADCredential = Get-Credential -Message "Enter AD domain credentials (e.g., vcaantech\youruser)" -ErrorAction Stop
                        if ($ADCredential) {
                            $ADCredential | Export-Clixml -Path $credPathAD -Force -ErrorAction Stop
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
                        # Use splatting for Get-ADComputer
                        $adComputerParams = @{
                            Filter     = "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'"
                            Properties = 'IPv4Address', 'OperatingSystem'
                            Server     = $config.InternalDomains.PrimaryDomain
                            Credential = $ADCredential
                            ErrorAction = 'Stop'
                        }
                        Get-ADComputer @adComputerParams |
                            Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                            Out-GridView -Title "Select Remote Desktop Server(s) to launch" -OutputMode Multiple -OutVariable SiteServers | Out-Null
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
                "81" {
                    # Launch WOOFware Reports Website
                    $ComputerNameStripped = Convert-VcaAu -AU $AU -Suffix ''
                    Start-Process "http://$ComputerNameStripped-db/reports/browse/WOOFware%20Reports"
                }
                "82" {
                    # Launch Fuse Website
                    $ComputerNameStripped = Convert-VcaAu -AU $AU -Suffix ''
                    Start-Process "https://$ComputerNameStripped-fuse:8443"
                }
                "83" {
                    Invoke-MenuOption83 -AU $AU -ADCredential $ADCredential
                }
                "h" {
                    Clear-Host
                    Write-Host "`n--- Main Menu for AU $AU (v$version) ---" -ForegroundColor Green
                    foreach ($key in ($menuOptions.Keys | Sort-Object @{Expression={ 
                        $match = [regex]::Match($_, '^(\d+)'); 
                        if ($match.Success) { [int]$match.Groups[1].Value } else { 999 }
                    }}, @{Expression={$_}})) {
                        Write-Host "$key. $($menuOptions[$key])" -ForegroundColor Cyan
                    }
                }
                "999" {
                    # New Session
                    Start-Process -FilePath "$PSScriptRoot\VCATechManager.cmd" -WorkingDirectory $PSScriptRoot
                }
                default {
                    Write-Host "Invalid choice. Please select 0-16, 19, 81-83, 999 or 'h' for menu." -ForegroundColor Red
                }
            }
        } while ($menuActive)
    }
} catch {
    Add-Content -Path $logPath -Value "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Error during script execution: $($_.Exception.Message)"
    Write-Host "An error occurred during script execution. Check the log file at $logPath for details." -ForegroundColor Red
}

# Reset console colors on exit (optional)
$host.UI.RawUI.BackgroundColor = "Black"
$host.UI.RawUI.ForegroundColor = "Gray"
Clear-Host
