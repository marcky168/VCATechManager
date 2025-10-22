#Harold.Kammermeyer@vca.com

function Get-MemoryUsage {
    Param (
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [string[]]$ComputerName,
        [pscredential]$Credential,
        [CimSession[]]$CimSession
    )
    begin {
        if (-not $CimSession) {
            foreach ($ComputerName_Item in $ComputerName) {
                try {
                    $CimSession = $CimSession + (New-CimSession -ComputerName $ComputerName_Item -Credential $Credential -ErrorAction Stop)
                }
                catch {
                    Write-Warning "[$ComputerName_Item] Memory Check: $($PSItem.Exception.Message)"
                }
            }
        }
    }
    process {
        foreach ($CimSession_Item in $CimSession) {
            try {
                $Response = Get-CimInstance -CimSession $CimSession_Item -ClassName Win32_OperatingSystem -Property TotalVisibleMemorySize, FreePhysicalMemory, LastBootUpTime -OperationTimeoutSec 10 -ErrorAction Stop
                [pscustomobject]@{
                    ComputerName    = $Response.PSComputerName
                    'UsedMemory(%)' = [decimal]('{0:N2}' -f (100 - (($Response.FreePhysicalMemory / $Response.TotalVisibleMemorySize) * 100)))
                    SysMemGB        = [decimal]('{0:N2}' -f ($Response.TotalVisibleMemorySize / 1MB))
                    BootUpTime      = $Response.LastBootUpTime
                }
            }
            catch {
                Write-Warning "[$($CimSession_Item.ComputerName)] Memory Check: $($PSItem.Exception.Message)"
            }
        }
    }
    end {
        if ($ComputerName -and $CimSession) { Remove-CimSession -CimSession $CimSession -ErrorAction Ignore }
    }
} #function