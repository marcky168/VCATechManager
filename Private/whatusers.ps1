#Harold.Kammermeyer@vca.com
#Requires -Version 3
#Requires -Modules ActiveDirectory

Import-Module -Name "$global:ScriptRoot\Private\lib\PoshRSJob\1.7.4.4\PoshRSJob.psm1"
Import-Module -Name "$global:ScriptRoot\Private\lib\PSTerminalServices"

function whatusers {
    param(
        [string[]]$ComputerName,
        [pscredential]$Credential
    )
    begin {
        # retrieve vdisk mounts from cluster NSs
        if ($Computername -match '-ns[0-9]|-ra[0-9]') {
            $Cluster = $true
            $DiskParams = @{
                ComputerName        = $ComputerName
                OperationTimeoutSec = 15
            }
            if ($Credential) { $DiskParams.Credential = $Credential }
            $DiskSession = New-CimSession @DiskParams #-ErrorAction SilentlyContinue
            $VDiskMountsRaw = Get-Disk -CimSession $DiskSession -ThrottleLimit 10 | Where-Object FriendlyName -eq 'Msft Virtual Disk'

            if ($VDiskMountsRaw) {
                $LDAPFilter = -join @($VDiskMountsRaw).foreach({ "(objectSID=$([regex]::Match($PSItem.Location, '(?i)S-\d-\d+-(\d+-){1,14}\d+').Value))" })
                $UserADList = Get-ADUser -LDAPFilter "(|$LDAPFilter)"

                $VDiskMounts = $VDiskMountsRaw | ForEach-Object {
                    Clear-Variable -Name UserADDisk, UserSID -ErrorAction Ignore

                    $UserSID = [regex]::Match($PSItem.Location, '(?i)S-\d-\d+-(\d+-){1,14}\d+').Value
                    $UserADDisk = @($UserADList).where({ $_.SID -eq $UserSID })

                    [PSCustomObject]@{
                        Computer       = $PSItem.PSComputerName
                        Name           = $UserADDisk.Name
                        SamAccountName = $UserADDisk.SamAccountName
                        DiskNumber     = $PSItem.DiskNumber
                        SID            = $UserSID
                        Location       = $PSItem.Location
                    }
                } #foreach
            }
        } # if cluster
    }
    process {
        $ComputerName | Start-RSJob -Name 'whatusersjobs' -Throttle 10 -ScriptBlock {
            param($ComputerName_Item)
            Import-Module -Name "$using:global:ScriptRoot\Private\lib\PSTerminalServices"

            try {
                Get-TSSession -ComputerName $ComputerName_Item -ErrorAction Stop |
                Select-Object -Property Server, UserName, ClientName, IPAddress, State, IdleTime, LoginTime, CurrentTime, SessionId
            }
            catch {
                Write-Warning "[$ComputerName_Item] $($PSItem.Exception.Message)"
            }
        } | Out-Null

        $TSSession = Get-RSJob -Name 'whatusersjobs' | Wait-RSJob -ShowProgress -Timeout 120 | Receive-RSJob

        if ($TSSession.UserName -ne '') {
            $LDAPFilter = -join ($TSSession.UserName.foreach({ if ($PSItem) { "(sAMAccountName=$PSItem)" } }))
            $UserADList = Get-ADUser -LDAPFilter "(|$LDAPFilter)" -Properties Department, Title, City, State, telephoneNumber, SID -ErrorAction Ignore
        }

        $TSSessionResults = foreach ($TSSession_Item in $TSSession) {
            Clear-Variable -Name UserAD -ErrorAction Ignore

            if ($Cluster) {
                # Collect user sessions to find any 'orphaned' UPDs (in next code segment)
                if (-not $TSSessionUsers) { [System.Collections.ArrayList]$TSSessionUsers = @() }
                if ($TSSession_Item.UserName) { $TSSessionUsers.Add($TSSession_Item.UserName) | Out-Null }
            }
            if ($TSSession_Item.UserName -and $TSSession_Item.UserName -ne 'webclock') {
                $UserAD = @($UserADList).where({ $PSItem.SamAccountName -eq $TSSession_Item.UserName })
            }
            [PSCustomObject]@{
                Computer    = $TSSession_Item.Server.ServerName
                SessionId   = $TSSession_Item.SessionId
                VDiskMount  = $(if ($VDiskMounts) { (@($VDiskMounts).Where( { $PSItem.SID -eq $UserAD.SID })).Computer })
                VDiskNumber = $(if ($VDiskMounts) { (@($VDiskMounts).Where( { $PSItem.SID -eq $UserAD.SID })).DiskNumber })
                UserName    = $TSSession_Item.UserName
                Title       = $(if ($UserAD.Title -like '*manager*' -or $UserAD.Title -like '*director*') { '*' + $UserAD.Title } else { $UserAD.Title })
                ClientName  = $TSSession_Item.ClientName
                IPAddress   = $TSSession_Item.IPAddress
                State       = $TSSession_Item.State
                IdleTime    = $(if ($TSSession_Item.IdleTime) { [timespan]('{0:dd}:{0:hh}:{0:mm}' -f $TSSession_Item.IdleTime) })
                LoginTime   = $TSSession_Item.LoginTime
                CurrentTime = $TSSession_Item.CurrentTime
                Department  = $UserAD.Department
                Location    = $(if ($UserAD.City) { $UserAD.City + ', ' + $UserAD.State })
                Phone       = $UserAD.telephoneNumber
                SID         = $UserAD.SID
            }
        } #foreach

        # Check cluster NSs for mounted UPDs (vhdx) with no associated user session
        if ($Cluster -and $VDiskMounts -and $TSSessionResults.UserName) {
            $VDiskResults = $VDiskMounts | Where-Object SamAccountName -NotIn $TSSessionUsers | ForEach-Object {
                Clear-Variable -Name UserAD -ErrorAction Ignore
                if ($PSItem.SamAccountName) {
                    $UserAD = Get-ADUser -Identity $PSItem.SamAccountName -Properties Department, Title, City, State, telephoneNumber -ErrorAction Ignore
                    [PSCustomObject]@{
                        Computer    = $PSItem.Computer
                        SessionId   = '99999'
                        VDiskMount  = $PSItem.Computer
                        VDiskNumber = $PSItem.DiskNumber
                        UserName    = $PSItem.SamAccountName
                        Title       = $(if ($UserAD.Title -like '*manager*' -or $UserAD.Title -like '*director*') { '*' + $UserAD.Title } else { $UserAD.Title })
                        ClientName  = 'User in logout process or'
                        IPAddress   = 'Potential orphan VHD'
                        State       = ''
                        IdleTime    = ''
                        LoginTime   = ''
                        CurrentTime = ''
                        Department  = $UserAD.Department
                        Location    = $(if ($UserAD.City) { $UserAD.City + ', ' + $UserAD.State })
                        Phone       = $UserAD.telephoneNumber
                        SID         = $PSItem.SID
                    }
                }
            }
            if (-not $VDiskResults) {
                [PSCustomObject]@{
                    Computer    = Convert-VcaAu -AU (@($ComputerName)[0]) -Prefix 'AU' -Suffix '' -NoLeadingZeros
                    SessionId   = ''
                    VDiskMount  = ''
                    VDiskNumber = ''
                    UserName    = ''
                    Title       = ''
                    ClientName  = 'No orphan disks found'
                    IPAddress   = ''
                    State       = ''
                    IdleTime    = ''
                    LoginTime   = ''
                    CurrentTime = ''
                    Department  = ''
                    Location    = ''
                    Phone       = ''
                    SID         = ''
                }
            }
            else {
                Write-Output $VDiskResults
            }
        }
        Write-Output $TSSessionResults
    }
    end {
        if ($DiskSession) { Remove-CimSession -CimSession $DiskSession -ErrorAction Ignore }
        Get-RSJob -Name 'whatusersjobs' | Remove-RSJob
    }
} #function