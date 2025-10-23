# VCATechManager AI Coding Instructions

## Project Overview
VCATechManager is a PowerShell-based IT management tool for VCA hospital networks. It provides a menu-driven interface for tasks like user session management, device connectivity checks, error logging, and infrastructure monitoring across multiple servers in an Active Directory environment.

## Architecture
- **Main Entry**: `VCATechManager.ps1` - Menu system that loads modules and calls functions.
- **Modules**: `Private/VCATechManagerFunctions.psm1` for shared utilities; `Private/lib/` for third-party modules (e.g., PoshRSJob, PSTerminalServices).
- **Scripts**: `Private/*.ps1` for specific functions (e.g., ADUserManagement.ps1).
- **Data**: `csv/` for hospital master Excel; `logs/` for timestamped log files.
- **Binaries**: `bin/` for tools like VNC viewer.
- **Credentials**: Stored in `Private/*.xml` (encrypted).

## Key Patterns & Conventions
- **Parallel Processing**: Use runspaces (via `[runspacefactory]::CreateRunspacePool`) for server queries to avoid blocking. Limit concurrency (e.g., max 10) to prevent network saturation. Example: `Get-UserSessionsParallel` in `VCATechManagerFunctions.psm1`.
- **Splatting**: Always use hashtables for long cmdlet parameters to improve readability. Example: `$invokeParams = @{ ComputerName = $server; ScriptBlock = $sb }; Invoke-Command @invokeParams`.
- **Logging & Error Handling**: Wrap operations in try-catch; use `Write-Log` for all actions; include stack traces in logs. Example: `try { Get-ADComputer @params } catch { Write-Log "Error: $($_.Exception.Message)" }`.
- **Caching**: Cache AD server lists and DNS resolutions with expiration (e.g., 10 minutes) using global hashtables to reduce queries.
- **User Interaction**: Use `Out-GridView` for selections; `Read-Host` for confirmations; `Write-Progress` for long operations.
- **Exports**: Use `Export-Results` helper for CSV exports to `reports/` folder with timestamps.
- **Credentials**: Load from XML early; validate with `Test-ADCredentials`; prompt if invalid.
- **Modularity**: Dot-source `Private/*.ps1`; import modules explicitly; export functions in PSM1 files.

## Developer Workflows
- **Build/Run**: Execute `VCATechManager.ps1` directly; no build step. Ensure RSAT and modules are installed.
- **Debugging**: Use `Write-Debug` for verbose output; check `logs/` for errors. Test functions in isolation (e.g., `Get-CachedServers -AU 123`).
- **Testing**: Manual testing per AU; verify parallel jobs complete without errors. Use mock data for unit-like tests.
- **Versioning & Updates**: Update version number in `Private/Version.txt` and append changes to `VCATechManager-Changelog.txt` with format "vX.Y - YYYY-MM-DD\n+ Change description". Script checks GitHub for updates at startup via `https://raw.githubusercontent.com/marcky168/VCATechManager/main/Private/Version.txt`.

## Integration Points
- **AD**: Query users/groups with `Get-ADUser/Get-ADGroupMember`; server discovery via `Get-ADComputer`.
- **DHCP/Networking**: Use `Get-DhcpServerv4Lease` for device MACs; `Get-NetNeighbor` for ARP.
- **Sessions**: `Get-TSSession` for RDP sessions; parse Security event logs (EventID 4624) for client IPs.
- **Email**: Outlook COM for drafts; embed signatures as base64.
- **VMware/SharePoint**: Connect via PnP.PowerShell for data pulls.
- **External Tools**: Launch VNC (`vncviewer.exe`), RDP (`mstsc.exe`), browsers for web interfaces.

## Specific Examples
- **Server Query**: `$servers = Get-CachedServers -AU $AU; Invoke-Command -ComputerName $servers -ScriptBlock { Get-WinEvent -LogName Application -FilterHashtable @{ProviderName='Woofware'; Level=2} }`
- **User Selection**: `$selectedUser = $adUsers | Out-GridView -Title "Select user" -OutputMode Single; if ($selectedUser) { User-LogonCheck -AU $AU -Username $selectedUser.SamAccountName }`
- **Parallel Job**: Use `Start-RSJob` for simple parallelism; prefer runspaces for complex async.

Focus on reliability, performance, and user experience. Avoid synchronous loops; always handle credential failures gracefully.