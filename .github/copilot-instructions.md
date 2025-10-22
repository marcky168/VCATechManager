# Copilot Instructions for VCATechManager

## Project Overview
This is a PowerShell-based IT automation suite for VCA hospital systems. The main script (`VCATechManager.ps1`) provides menu-driven tools for managing hospital administrative units (AUs), including server queries, user logons, device connectivity, error diagnostics, DHCP management, and integrations with ServiceNow, vSphere, and more.

## Architecture
- **Main Script**: `VCATechManager.ps1` loads functions from `Private/` folder and custom module `VCATechManagerFunctions.psm1`
- **Modules**: External dependencies in `Private/lib/` (ActiveDirectory, PoshRSJob, PSTerminalServices, ImportExcel, PnP.PowerShell, Autoload)
- **Data Flow**: Queries Active Directory for servers/users, DHCP for leases/reservations, Windows Event Logs for errors; caches server lists and IP resolutions in memory hashes; loads hospital master from SharePoint
- **Output**: Logs to `logs/`, exports CSVs to `reports/`, displays via Write-Host with colors; launches external tools like VNC, RDP, browsers

## Key Patterns
- **Function Naming**: Use `Verb-Noun` format (e.g., `Abaxis-MacAddressSearch`, `Woofware-ErrorsCheck`, `Get-UserSessionsParallel`)
- **AU Parameter**: All major functions take `-AU` parameter for hospital unit (3-6 digits); validate with regex `^\d{3,6}$`
- **Parallel Processing**: Use `Start-RSJob` for concurrent server queries with `Wait-RSJob | Receive-RSJob | Remove-RSJob`; use `PSSessionOption` for timeouts
- **Caching**: Store server lists in `$validAUs` hash with timestamps; IP resolutions in `$ipCache` with expiration
- **Error Handling**: Wrap operations in `try-catch`, log with `Write-Log` including stack traces; use `-ErrorAction Stop` for propagation
- **Export Helper**: Use `Export-Results` function for CSV exports with user confirmation and progress bars
- **Logging**: Always call `Write-Log` for actions; `Write-ConditionalLog` for verbose mode; toggle with `$verboseLogging`
- **Progress Bars**: Use `Write-Progress` with `Activity`, `Status`, and `PercentComplete` for long operations
- **User Interaction**: Use `Read-Host` for inputs; `Out-GridView` for selections; menu with hashtable for options
- **MAC Address Handling**: Normalize MACs with `$mac.Replace("-", "").ToUpper()` for DHCP comparisons; define device groups by prefixes
- **Email Creation**: Use Outlook COM objects to create drafts with HTMLBody and embedded base64 images
- **Time Zone Conversion**: Map IANA to Windows TZ IDs (e.g., "America/Chicago" to "Central Standard Time") for hospital time display
- **DHCP Splatting**: Use hashtables for `Get-DhcpServerv4Lease`, `Add-DhcpServerv4Reservation`
- **vSphere URLs**: Determine based on time zones (East/Central vs West Coast)
- **Session Management**: Use `Get-TSSession` for active sessions; fetch ClientIP from Security event logs with XPath
- **Module Loading**: Import custom modules like `VCATechManagerFunctions.psm1`; auto-install PoshRSJob locally if missing

## Developer Workflows
- **Running**: Execute `VCATechManager.ps1` in PowerShell; enter AU number, select from 19 menu options
- **Debugging**: Use `Write-Host` with colors for UI; check `logs/VCATechManager_log_*.txt` for details; use `Write-Debug` in jobs
- **Adding Functions**: Create in `Private/` folder, dot-source in main script or add to `VCATechManagerFunctions.psm1`, add to menu hashtable
- **Testing**: Run functions individually with test AU (e.g., 966), verify outputs, logs, and error handling
- **Version Checking**: Query GitHub raw URL at startup; compare with `$version` variable
- **Credential Management**: Store AD creds with `Export-Clixml` to `Private/vcaadcred.xml`; validate with `Test-ADCredentials`
- **Versioning**: Increment `$version` (e.g., 1.6 to 1.7) for changes; semantic versioning (MAJOR.MINOR.PATCH)
- **Changelog**: Update `VCATechManager-Changelog.txt` with dated entries; newest at top, oldest at bottom; only after testing

## Integration Points
- **Active Directory**: Query users/servers with `Get-ADComputer`, `Get-ADUser`; validate creds with `Test-ADCredentials`
- **DHCP**: Manage leases/reservations via `Get-DhcpServerv4Lease`, `Add-DhcpServerv4Reservation` on `phhospdhcp2.vcaantech.com`
- **SharePoint**: Download hospital master via PnP.PowerShell from `https://vca365.sharepoint.com`; load with `Import-Excel`
- **ServiceNow**: Launch URLs like `https://marsvh.service-now.com/now/nav/ui/classic/params/target/incident_list.do?sysparm_query=u_departmentLIKE$AU`
- **External Tools**: Launch VNC (`$PSScriptRoot\Private\bin\vncviewer.exe`), RDP (`mstsc`), browser (`msedge`); Angry IP Scanner on scopes
- **Outlook**: Automate email drafts with COM objects for error reporting; HTMLBody with base64 images
- **vSphere**: Open management URLs based on time zones (East/Central: one URL, West: another)
- **Hospital APIs**: Fetch hours from `uat.vcahospitals.com` or `uat.vcacanada.com` APIs; display with `Format-List`
- **Event Logs**: Query Security logs with XPath for ClientIP; Application logs for Woofware errors

## Conventions
- **Credentials**: Store AD creds with `Export-Clixml` to `Private/vcaadcred.xml`; validate with `Test-ADCredentials`
- **Hospital Master**: Load from `Private/csv/HOSPITALMASTER.xlsx` using `Import-Excel`; update from SharePoint with `Update-HospitalMaster`
- **Time Zones**: Convert IANA to Windows TZ IDs (e.g., "America/Chicago" to "Central Standard Time") for hospital time display
- **Progress Bars**: Use `Write-Progress` with `Activity`, `Status`, `PercentComplete` for long operations
- **User Interaction**: Use `Read-Host` for inputs; `Out-GridView` for selections; menu with hashtable sorted numerically
- **AU Formatting**: Use `Convert-VcaAU` function for hostname generation with prefixes/suffixes (e.g., `-Ilo`, `-EsxiHost`)
- **Group Naming**: AD groups named as 'H' + AU.PadLeft(4, '0') (e.g., H0966)
- **Server Filtering**: Exclude CNF duplicates with `-and Name -notlike '*CNF:*'`; filter by `OperatingSystem -like '*Server*'`
- **Event Log Queries**: Use `Get-WinEvent` with `-FilterHashtable` for efficient log searches; XPath for Security logs
- **Ping Tests**: Use custom `Get-PingStatus` filter; `Test-Connection` with `-Count` and `-Quiet`
- **MAC Normalization**: `$mac.Replace("-", "").ToUpper()` for consistent formatting; group by prefixes (VS2, HM5, etc.)
- **Fuse Types**: Determine Virtual vs Physical based on IP (10.242* = Virtual)
- **Menu Sorting**: Sort keys numerically, handling non-numeric suffixes with regex

## Examples
- **Server Query**: `$servers = Get-ADComputer -Filter "Name -like '$SiteAU-ns*'" -Properties Name | Select-Object -ExpandProperty Name | Sort-Object Name`
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