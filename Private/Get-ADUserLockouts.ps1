# Modified by Harold for remote use
# usage example  get-aduser hailey.serino | Get-ADUserLockouts -StartTime (get-date).AddDays(-1) -EndTime (get-date)
# from  https://theposhwolf.com/howtos/Get-ADUserLockouts/

Function Get-ADUserLockouts {
    [CmdletBinding(
        DefaultParameterSetName = 'All'
    )]
    param (
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'ByUser'
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]$Identity,
        [datetime]$StartTime = (Get-Date).AddDays(-1),
        [datetime]$EndTime = (Get-Date),
        [pscredential]$Credential
    )
    Begin {
        $filterHt = @{
            LogName = 'Security'
            ID      = 4740
        }
        if ($PSBoundParameters.ContainsKey('StartTime')) {
            $filterHt['StartTime'] = $StartTime
        }
        if ($PSBoundParameters.ContainsKey('EndTime')) {
            $filterHt['EndTime'] = $EndTime
        }
        $PDCEmulator = (Get-ADDomain).PDCEmulator
        # Query the event log just once instead of for each user if using the pipeline
        $events = Invoke-Command -ComputerName $PDCEmulator -ScriptBlock { Get-WinEvent -FilterHashtable $using:filterHt | Select-Object -Property TimeCreated, Properties } -Credential $Credential
    }
    Process {
        if ($PSCmdlet.ParameterSetName -eq 'ByUser') {
            $user = Get-ADUser $Identity
            # Filter the events
            $output = $events | Where-Object { $_.Properties[0].Value -eq $user.SamAccountName }
        }
        else {
            $output = $events
        }
        foreach ($event in $output) {
            [pscustomobject]@{
                UserName       = $event.Properties[0].Value
                CallerComputer = $event.Properties[1].Value
                TimeStamp      = $event.TimeCreated
            }
        }
    }
    End {}
}

## Usage Example
##  get-aduser kishore.reddy | Get-ADUserLockouts -StartTime (get-date).AddDays(-1) -EndTime (get-date)