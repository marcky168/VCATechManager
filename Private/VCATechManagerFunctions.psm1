# Custom module for VCATechManager shared functions
# This module contains reusable functions for session management and other utilities.

# Function: Get-UserSessionsParallel
# Description: Retrieves user sessions in parallel across servers for a given AU and username using runspaces for efficiency.
# Parameters:
#   - AU: Administrative Unit number (string)
#   - Username: Username to query sessions for (string)
# Returns: Array of custom objects with session details
function Get-UserSessionsParallel {
    param([string]$AU, [string]$Username)

    # Log the start of the operation
    Write-Log "Starting Get-UserSessionsParallel for AU $AU, User $Username"

    try {
        # Fetch cached servers for the AU
        $servers = Get-CachedServers -AU $AU
    } catch {
        # Handle errors in server fetching
        Write-Host "Error fetching servers for AU $AU : $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error in Get-UserSessionsParallel: $($_.Exception.Message)"
        return @()
    }

    # Use runspaces for parallel processing (more efficient than jobs)
    $maxConcurrency = 10  # Limit concurrency to prevent network saturation
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxConcurrency)
    $runspacePool.Open()

    $totalServers = $servers.Count
    $i = 0
    $handles = @()
    $powershellInstances = @()

    # Loop through each server and start a runspace
    foreach ($server in $servers) {
        $i++
        Write-Progress -Activity "Querying user sessions" -Status "Server $i of $totalServers : $server" -PercentComplete (($i / $totalServers) * 100)

        # Create PowerShell instance for runspace
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool

        # Define the scriptblock for session querying
        $scriptBlock = {
            param($server, $Username)
            try {
                # Configure session options for reliability
                $sessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000

                # Use splatting for Invoke-Command
                $invokeParams = @{
                    ComputerName  = $server
                    SessionOption = $sessionOption
                    ScriptBlock   = {
                        param($Username)
                        # Import PSTerminalServices module for session queries
                        Import-Module -Name "$using:PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue

                        # Retrieve active or disconnected sessions for the user
                        $sessions = Get-TSSession -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Active' -or $_.State -eq 'Disconnected' } | Where-Object { $_.UserName -eq $Username }

                        # Initialize results array
                        $results = @()
                        foreach ($session in $sessions) {
                            # Attempt to fetch ClientIP from event logs
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

                            # Create custom object for each session
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
                    }
                    ArgumentList = $Username
                }
                Invoke-Command @invokeParams
            } catch {
                Write-Debug "Error querying sessions on $server : $($_.Exception.Message)"
                # Return error object if query fails
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

        # Add script and arguments to PowerShell instance
        $ps.AddScript($scriptBlock).AddArgument($server).AddArgument($Username)

        # Begin invoke asynchronously
        $handle = $ps.BeginInvoke()
        $handles += $handle
        $powershellInstances += $ps
    }

    # Wait for all runspaces to complete
    while ($handles | Where-Object { -not $_.IsCompleted }) {
        Start-Sleep -Milliseconds 100
    }

    # Collect results and errors
    $results = @()
    for ($j = 0; $j -lt $powershellInstances.Count; $j++) {
        $ps = $powershellInstances[$j]
        $handle = $handles[$j]
        try {
            $result = $ps.EndInvoke($handle)
            if ($result) {
                $results += $result
            }
        } catch {
            Write-Debug "Error in runspace for server $($servers[$j]): $($_.Exception.Message)"
        }
        # Collect errors from streams
        if ($ps.Streams.Error) {
            foreach ($error in $ps.Streams.Error) {
                Write-Debug "Runspace error: $($error.Exception.Message)"
            }
        }
        $ps.Dispose()
    }

    # Close runspace pool
    $runspacePool.Close()
    $runspacePool.Dispose()

    Write-Progress -Activity "Querying user sessions" -Completed

    # Log completion and return results
    Write-Log "Get-UserSessionsParallel completed: Found $($results.Count) sessions"
    Write-Debug "Session details: $($results | Out-String)"
    return $results
}

# Export the function for use in importing scripts
Export-ModuleMember -Function Get-UserSessionsParallel
