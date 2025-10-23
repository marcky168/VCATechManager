# Copilot Instructions for VCATechManager

## Project Overview
This is a PowerShell-based IT automation suite for VCA hospital systems. The main script (`VCATechManager.ps1`) provides menu-driven tools for managing hospital administrative units (AUs), including server queries, user logons, device connectivity, error diagnostics, DHCP management, and integrations with ServiceNow, vSphere, and more.

## Architecture
- **Main Script**: `VCATechManager.ps1` loads functions from `Private/` folder and custom module `VCATechManagerFunctions.psm1`
- **Modules**: External dependencies in `Private/lib/` (ActiveDirectory, PoshRSJob, PSTerminalServices, ImportExcel, PnP.PowerShell, Autoload, VMware modules)
- **Data Flow**: Queries Active Directory for servers/users, DHCP for leases/reservations, Windows Event Logs for errors; caches server lists and IP resolutions in memory hashes; loads hospital master from SharePoint/Excel
- **Output**: Logs to `logs/`, exports CSVs to `reports/`, displays via Write-Host with colors; launches external tools like VNC, RDP, browsers

## Key Patterns
- **Function Naming**: Use `Verb-Noun` format (e.g., `Abaxis-MacAddressSearch`, `Woofware-ErrorsCheck`, `Get-UserSessionsParallel`)
- **AU Parameter**: All major functions take `-AU` parameter for hospital unit (3-6 digits); validate with regex `^\d{3,6}$`
- **Parallel Processing**: Use runspaces for concurrent operations (more efficient than jobs); limit concurrency (e.g., max 10) to prevent network saturation; fallback to `Start-RSJob` with `Wait-RSJob | Receive-RSJob | Remove-RSJob` for compatibility
- **Caching**: Store server lists in `$validAUs` hash with timestamps (10min expiry); IP resolutions in `$ipCache` with expiration
- **Error Handling**: Wrap operations in `try-catch`, log with `Write-Log` including stack traces; use `-ErrorAction Stop` for propagation; validate credentials with `Test-ADCredentials`
- **Export Helper**: Use `Export-Results` function for CSV exports with user confirmation and progress bars
- **Logging**: Always call `Write-Log` for actions; `Write-ConditionalLog` for verbose mode; toggle with `$verboseLogging`; logs include timestamps and stack traces
- **Progress Bars**: Use `Write-Progress` with `Activity`, `Status`, and `PercentComplete` for long operations
- **User Interaction**: Use `Read-Host` for inputs; `Out-GridView` for selections; menu with hashtable sorted numerically (regex for non-numeric keys like "2b")
- **MAC Address Handling**: Normalize MACs with `$mac.Replace("-", "").ToUpper()` for DHCP comparisons; define device groups by prefixes (VS2, HM5, CBC)
- **Email Creation**: Use Outlook COM objects to create drafts with HTMLBody and embedded base64 images; ensure Outlook is running with `Start-OutlookIfNeeded`
- **Time Zone Conversion**: Map IANA to Windows TZ IDs (e.g., "America/Chicago" to "Central Standard Time") for hospital time display in prompts
- **DHCP Splatting**: Use hashtables for `Get-DhcpServerv4Lease`, `Add-DhcpServerv4Reservation` on `phhospdhcp2.vcaantech.com`
- **vSphere URLs**: Determine based on time zones (East/Central vs West Coast); use VMware modules for integration
- **Session Management**: Use `Get-TSSession` for active sessions; fetch ClientIP from Security event logs with XPath; support VNC/Shadow to client IPs
- **Module Loading**: Import custom modules like `VCATechManagerFunctions.psm1`; auto-install PoshRSJob locally if missing; suppress PnP update checks
- **Splatting**: Use hashtables for long parameter lists (e.g., `Get-ADComputer`, `Invoke-Command`) for readability
- **Version Checking**: Query GitHub raw URL at startup; compare with `$version` variable for updates

## Developer Workflows
- **Running**: Execute `VCATechManager.ps1` in PowerShell; enter AU number (3-6 digits), select from menu options (19+ including sub-options like 2b, 14u)
- **Debugging**: Use `Write-Host` with colors for UI; check `logs/VCATechManager_log_*.txt` for details; use `Write-Debug` in runspaces; verify with test AU (e.g., 966)
- **Adding Functions**: Create in `Private/` folder, dot-source in main script or add to `VCATechManagerFunctions.psm1`, add to menu hashtable with numeric sorting
- **Testing**: Run functions individually with test AU, verify outputs, logs, error handling, and parallel operations; test credential validation and module loading
- **Version Checking**: Query GitHub raw URL at startup (e.g., `https://raw.githubusercontent.com/yourrepo/marc-tools/main/version.txt`); compare with `$version` variable
- **Credential Management**: Store AD creds with `Export-Clixml` to `Private/vcaadcred.xml`; validate with `Test-ADCredentials`; prompt if invalid
- **Versioning**: Increment `$version` (e.g., 1.10 to 1.11) for changes; semantic versioning (MAJOR.MINOR.PATCH); update lastWritten timestamp
- **Changelog**: Update `VCATechManager-Changelog.txt` with dated entries; newest at top, oldest at bottom; include rationale and testing notes; only after testing
- **Menu Updates**: Add options to `$menuOptions` hashtable; sort numerically with regex for suffixes; handle new sessions with option 999

## Integration Points
- **Active Directory**: Query users/servers with `Get-ADComputer`, `Get-ADUser`; validate creds with `Test-ADCredentials`; filter servers with `-and Name -notlike '*CNF:*'` and `OperatingSystem -like '*Server*'`
- **DHCP**: Manage leases/reservations via `Get-DhcpServerv4Lease`, `Add-DhcpServerv4Reservation` on `phhospdhcp2.vcaantech.com`; normalize MACs for comparisons
- **SharePoint**: Download hospital master via PnP.PowerShell from `https://vca365.sharepoint.com`; load with `Import-Excel`; update with `Update-HospitalMaster`
- **ServiceNow**: Launch URLs like `https://marsvh.service-now.com/now/nav/ui/classic/params/target/incident_list.do?sysparm_query=u_departmentLIKE$AU`
- **External Tools**: Launch VNC (`$PSScriptRoot\Private\bin\vncviewer.exe`), RDP (`mstsc`), browser (`msedge`); Angry IP Scanner on scopes; Outlook for emails
- **Outlook**: Automate email drafts with COM objects for error reporting; HTMLBody with base64 images; ensure running with `Start-OutlookIfNeeded`
- **vSphere**: Open management URLs based on time zones (East/Central: one URL, West: another); use VMware modules for integration
- **Hospital APIs**: Fetch hours from `uat.vcahospitals.com` or `uat.vcacanada.com` APIs; display with `Format-List`; handle time zone conversions
- **Event Logs**: Query Security logs with XPath for ClientIP; Application logs for Woofware errors; use `Get-WinEvent` with `-FilterHashtable`
- **VMware Modules**: Import VMware.VimAutomation.* for vSphere operations; determine URLs by time zone mappings

## Conventions
- **Credentials**: Store AD creds with `Export-Clixml` to `Private/vcaadcred.xml`; validate with `Test-ADCredentials`; prompt if invalid
- **Hospital Master**: Load from `Private/csv/HOSPITALMASTER.xlsx` using `Import-Excel`; update from SharePoint with `Update-HospitalMaster`
- **Time Zones**: Convert IANA to Windows TZ IDs (e.g., "America/Chicago" to "Central Standard Time") for hospital time display in prompts
- **Progress Bars**: Use `Write-Progress` with `Activity`, `Status`, `PercentComplete` for long operations
- **User Interaction**: Use `Read-Host` for inputs; `Out-GridView` for selections; menu with hashtable sorted numerically (regex for non-numeric keys like "2b", "14u")
- **AU Formatting**: Use `Convert-VcaAU` function for hostname generation with prefixes/suffixes (e.g., `-Ilo`, `-EsxiHost`); supports `-FQDN`, `-Clipboard`, `-RetainSuffixNumber`
- **Group Naming**: AD groups named as 'H' + AU.PadLeft(4, '0') (e.g., H0966)
- **Server Filtering**: Exclude CNF duplicates with `-and Name -notlike '*CNF:*'`; filter by `OperatingSystem -like '*Server*'`; include Util servers
- **Event Log Queries**: Use `Get-WinEvent` with `-FilterHashtable` for efficient log searches; XPath for Security logs (ClientIP extraction)
- **Ping Tests**: Use custom `Get-PingStatus` filter; `Test-Connection` with `-Count` and `-Quiet`; ping sweeps for ARP population
- **MAC Normalization**: `$mac.Replace("-", "").ToUpper()` for consistent formatting; group by prefixes (VS2: "00-07-32", HM5, CBC)
- **Fuse Types**: Determine Virtual vs Physical based on IP (10.242* = Virtual)
- **Menu Sorting**: Sort keys numerically, handling non-numeric suffixes with regex; handle new sessions (option 999)
- **Parameter Validation**: Validate AU with regex `^\d{3,6}$`; use `-ErrorAction Stop` for propagation

## Examples
- **Server Query**: `$servers = Get-ADComputer -Filter "Name -like '$SiteAU-ns*'" -Properties Name | Select-Object -ExpandProperty Name | Sort-Object Name`
- **Parallel Runspace**: `$runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrency); $ps = [powershell]::Create(); $ps.RunspacePool = $runspacePool`
- **Parallel Job**: `$jobs += Start-RSJob -ScriptBlock { param($server) ... } -ArgumentList $server; $results = $jobs | Wait-RSJob | Receive-RSJob | Remove-RSJob`
- **DHCP Splatting**: `$leaseParams = @{ ComputerName = $dhcpServer; ScopeId = $scopeId }; Get-DhcpServerv4Lease @leaseParams`
- **MAC Normalization**: `$mac.Replace("-", "").ToUpper()` for consistent formatting; group by prefixes like VS2: "00-07-32"
- **Outlook Email Draft**: `$outlook = New-Object -ComObject Outlook.Application; $mail = $outlook.CreateItem(0); $mail.HTMLBody = "<html>...</html>"; $mail.Save()`
- **vSphere URL Selection**: `if ($timeZone -like "*Chicago*" -or "*New_York*") { $url = "east-vsphere.com" } else { $url = "west-vsphere.com" }`
- **AU Conversion**: `Convert-VcaAU -AU $AU -Suffix '-gw'` for gateway hostnames; `Convert-VcaAU -AU $AU -Ilo` for ILO interfaces
- **Event Log Search**: `Get-WinEvent -FilterHashtable @{logname='Application';ProviderName='Woofware'; level=2 ;id=100,101,102} -MaxEvents 50`
- **ClientIP from Logs**: `$filterXPath = "*[System[EventID=4624] and EventData/Data[@Name='TargetUserName']='$username']"; $event = Get-WinEvent -LogName Security -FilterXPath $filterXPath -MaxEvents 1`
- **Hospital Hours API**: `Invoke-RestMethod -Uri "https://uat.vcahospitals.com/api/content/hospital/getUSHospitalHours?HospitalID=$id" | Format-List`

## Key Files
- `VCATechManager.ps1`: Main script
- `Private/`: Function definitions (e.g., `ADUserManagement.ps1`)
- `Private/lib/`: External modules
- `Private/csv/HOSPITALMASTER.xlsx`: Hospital data
- `logs/`: Execution logs
- `reports/`: CSV exports
- `VCATechManager-Changelog.txt`: Version history and change log
- `VCATechManagerFunctions.psm1`: Custom module for shared functions