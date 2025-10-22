# UserSessions Module: Contains functions for user session management
# Version: 1.0
# Author: Grok

# Updated: Enhanced for consistency with requirements: added Write-Progress, refined ClientIP fetching, try-catch, logging, splatting, and date format 'yyyy-MM-dd HH:mm:ss'
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
        # Splatting for Start-RSJob
        $jobParams = @{
            Name         = $server
            ScriptBlock  = {
                param($server, $Username)
                try {
                    # Splatting for Invoke-Command
                    $invokeParams = @{
                        ComputerName  = $server
                        SessionOption = New-PSSessionOption -OperationTimeout 60000 -IdleTimeout 60000
                        ScriptBlock   = {
                            param($Username)
                            Import-Module -Name "$using:PSScriptRoot\Private\lib\PSTerminalServices" -ErrorAction SilentlyContinue
                            $sessions = Get-TSSession -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Active' -or $_.State -eq 'Disconnected' } | Where-Object { $_.UserName -eq $Username }
                            $results = @()
                            foreach ($session in $sessions) {
                                $clientIP = "N/A"
                                try {
                                    $escapedUsername = $Username -replace "'", "''"
                                    $filterXPath = "*[System[EventID=4624] and EventData/Data[@Name='TargetUserName']='$escapedUsername']"
                                    # Splatting for Get-WinEvent
                                    $eventParams = @{
                                        LogName      = 'Security'
                                        FilterXPath  = $filterXPath
                                        MaxEvents    = 1
                                        ErrorAction  = 'Stop'
                                    }
                                    $event = Get-WinEvent @eventParams | Select-Object -First 1
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
                        }
                        ArgumentList = $Username
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

# Export the function for use in importing scripts
Export-ModuleMember -Function Get-UserSessionsParallel
