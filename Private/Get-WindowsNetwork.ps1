#Harold.Kammermeyer@vca.com
#Requires -Version 3
#Requires -Modules PoshRSJob
#Version: 230526
## Get Windows Network Adapter

function Get-WindowsNetwork {
    [CmdletBinding()]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('Name')]
        [ValidateNotNullOrEmpty()]
        [String[]]$ComputerName = $env:ComputerName,
        [PSCredential]$Credential
    )

    begin {
        # Remove empty lines
        $ComputerName = $ComputerName | Where-Object { $PSItem }
        # Remove duplicates
        $ComputerName = $ComputerName.ToLower() | Select-Object -Unique
    }
    process {
        $ComputerName | Start-RSJob -Name 'WindowsNetworkJob' -VariablesToImport Credential -Throttle 64 -ScriptBlock {
            $ComputerName_Item = $_
            try {
                $Session = New-CimSession -ComputerName $ComputerName_Item -Credential $Credential -OperationTimeoutSec 10 -ErrorAction Stop

                $CimParams = @{
                    CimSession          = $Session
                    ClassName           = 'Win32_NetworkAdapterConfiguration'
                    Property            = @(
                        'Index'
                        'Description'
                        'DNSHostName'
                        'DNSServerSearchOrder'
                        'IPAddress'
                        'DefaultIPGateway'
                        'IPSubnet'
                        'DHCPEnabled'
                        'MACAddress'
                    )
                    Filter              = 'IPEnabled=True'
                    OperationTimeoutSec = 10
                }
                $Response = Get-CimInstance @CimParams

                $CimParams2 = @{
                    CimSession          = $Session
                    ClassName           = 'Win32_OperatingSystem'
                    Property            = 'Caption'
                    OperationTimeoutSec = 10
                }
                $Response2 = Get-CimInstance @CimParams2
            }
            catch {
                $ErrorMessage = $PSItem.Exception.Message
            }
            $Response | ForEach-Object {
                [pscustomobject]@{
                    Name             = $ComputerName_Item
                    OperatingSystem  = $Response2.Caption
                    Index            = $PSItem.Index
                    Description      = $PSItem.Description
                    DNS              = $PSItem.DNSServerSearchOrder -join ", "
                    IP1              = @($PSItem.IPAddress)[0]
                    IP2              = @($PSItem.IPAddress)[1]
                    DefaultIPGateway = $PSItem | Select-Object -ExpandProperty DefaultIPGateway
                    IPSubnet         = $PSItem.IPSubnet -join ", "
                    DHCPEnabled      = $PSItem.DHCPEnabled
                    MACAddress       = $PSItem.MACAddress
                    Error            = $ErrorMessage
                }
            }
        } | Out-Null #Start-RSJob ScriptBlock
    } #process
    end {
        Get-RSJob -Name 'WindowsNetworkJob' | Wait-RSJob -ShowProgress -Timeout 600 | Receive-RSJob
        Get-RSJob -Name 'WindowsNetworkJob' | Remove-RSJob
    }
} #function