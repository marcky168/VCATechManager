#Harold.Kammermeyer@vca.com
#Requires -Version 3
#Requires -Modules ActiveDirectory

#Import-Module "$PSScriptRoot\lib\PSTerminalServices"

function whatdisk {
    param(
        [string[]]$ComputerName
    )
    if ($Computername -match '-ns[0-9]') {
        $Session = New-CimSession -ComputerName $ComputerName -ErrorAction SilentlyContinue
        $VDiskMounts = Get-Disk -CimSession $Session | Where-Object FriendlyName -eq 'Msft Virtual Disk' | 
            ForEach-Object {
                $PSItem | Select-Object -Property PSComputerName,
                @{n='Name';e={(Get-ADUser -Filter "SID -like '$($PSItem.Location -match "S-\d-\d+-(\d+-){1,14}\d+" | Out-Null; $Matches[0])'").Name}},
                @{n='SamAccountName';e={(Get-ADUser -Filter "SID -like '$($PSItem.Location -match "S-\d-\d+-(\d+-){1,14}\d+" | Out-Null; $Matches[0])'").SamAccountName}},
                DiskNumber,
                @{n='SID';e={$PSItem.Location -match "S-\d-\d+-(\d+-){1,14}\d+" | Out-Null; $Matches[0]}}, Location
            }
        $VDiskMounts
    }
   <# foreach ($Computer in $ComputerName) {
        try {
            #$TSSession = Get-TSSession -ComputerName $Computer -ErrorAction Stop | Select-Object UserName, ClientName, IPAddress, State, IdleTime, LoginTime, CurrentTime, SessionId
            #foreach ($ts in $TSSession) {
                $userAD = $(if ($ts.UserName -and $ts.UserName -ne 'webclock') { (Get-ADUser -Identity $ts.UserName -Properties Department, Title, City, State, telephoneNumber, SID -ErrorAction SilentlyContinue) })
                [PSCustomObject]@{
                    Computer    = $Computer
               #    SessionId   = $ts.SessionId
               #    UserName    = $ts.UserName
                    Title       = $(if ($userAD.Title -like '*manager*' -or $userAD.Title -like '*director*') { '*' + $userAD.Title } else { $userAD.Title })
               #    ClientName  = $ts.ClientName
               #    IPAddress   = $ts.IPAddress
               #    State       = $ts.State
               #    IdleTime    = [timespan]('{0:hh}:{0:mm}' -f $ts.IdleTime)
               #    LoginTime   = $ts.LoginTime
               #    LocalTime   = $ts.CurrentTime
               #    Department  = $userAD.Department
                    Location    = $(if ($userAD.City) { $userAD.City + ', ' + $userAD.State })
               #    Phone       = $userAD.telephoneNumber
                    SID         = $userAD.SID
                    VDiskMount  = ($VDiskMounts.Where({$PSItem.SID -eq $userAD.SID})).PSComputerName
                    VDiskNumber = ($VDiskMounts.Where({$PSItem.SID -eq $userAD.SID})).DiskNumber
                }
            } #foreach
        } #foreach
        catch {
            #intentionally left blank
        }
    } #>
    if ($Session) { Remove-CimSession -CimSession $Session -ErrorAction SilentlyContinue }
} #function