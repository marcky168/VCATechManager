#Harold.Kammermeyer@vca.com
## Get RDS Connection settings for Server 2012+.

Import-Module -Name RemoteDesktop -WarningAction Ignore 3>$null

function Get-RDSConnectionConfig {
    [CmdletBinding()]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('Name')]
        [String[]]
        $ComputerName = $env:ComputerName,

        [PSCredential]
        $Credential
    )

    begin {
        #intentionally left empty
    }
    process {
        $ComputerName | Start-RSJob -Name 'RdsJob' -Throttle 64 -ModulesToImport RemoteDesktop -ScriptBlock {
            try {
                <#
                $ConnectionBroker = $(
                    if ($_ -match '-ns[0-9]{1,2}$') {
                        $_ -replace $Matches.Values, '-fs.vcaantech.com' }
                    else {
                        "$_.vcaantech.com"
                    }
                )
                #>
                <#
                $CollectionName = $(
                    if ($_ -match '-ns[0-9]{1,2}$|-fs[0-9]{0,2}$') {
                        $_ -replace $Matches.Values, '-nslb'
                    }
                    else {
                        $_
                    }
                )
                #>
                $ConnectionBroker = [System.Net.Dns]::GetHostEntry($_).HostName
                $CollectionName = (Get-RDSessionCollection -ConnectionBroker $ConnectionBroker).CollectionName

                $RdsParams = @{
                    ConnectionBroker = $ConnectionBroker
                    CollectionName   = $CollectionName
                    Connection       = $true
                }
                $Response = Get-RDSessionCollectionConfiguration @RdsParams -ErrorAction Stop
            }
            catch {
                $ErrorMessage = $_.Exception.Message
            }
            [pscustomobject]@{
                ComputerName                  = $_
                ConnectionBroker              = $ConnectionBroker
                CollectionName                = $Response.CollectionName
                DisconnectedSessionLimitMin   = $Response.DisconnectedSessionLimitMin
                BrokenConnectionAction        = $Response.BrokenConnectionAction
                TemporaryFoldersDeletedOnExit = $Response.TemporaryFoldersDeletedOnExit
                AutomaticReconnectionEnabled  = $Response.AutomaticReconnectionEnabled
                ActiveSessionLimitMin         = $Response.ActiveSessionLimitMin
                IdleSessionLimitMin           = $Response.IdleSessionLimitMin
                Error                         = $ErrorMessage
            }
        } | Out-Null #Start-RSJob ScriptBlock
    } #process
    end {
        Get-RSJob -Name 'RdsJob' | Wait-RSJob -Timeout 600 | Receive-RSJob
        Get-RSJob -Name 'RdsJob' | Remove-RSJob
    }
} #function