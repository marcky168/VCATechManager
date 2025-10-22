#Harold.Kammermeyer@vca.com

function Get-DiskUsage {
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
                    Write-Warning "[$ComputerName_Item] Disk Check: $($PSItem.Exception.Message)"
                }
            }
        }
    }
    process {
        foreach ($CimSession_Item in $CimSession) {
            try {
                $Response = Get-CimInstance -CimSession $CimSession_Item -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -Property Name, FileSystem, FreeSpace, Size -OperationTimeoutSec 10 -ErrorAction Stop

                foreach ($LogicalDisk_Item in $Response) {
                    [pscustomobject]@{
                        ComputerName   = $LogicalDisk_Item.PSComputerName
                        Name           = $LogicalDisk_Item.Name
                        FileSystem     = $LogicalDisk_Item.FileSystem
                        FreeGB         = [decimal]('{0:N2}' -f ($LogicalDisk_Item.FreeSpace / 1GB))
                        'FreeSpace(%)' = [decimal]('{0:N2}' -f ((($LogicalDisk_Item.FreeSpace) / ($LogicalDisk_Item.Size)) * 100))
                        CapacityGB     = [decimal]('{0:N2}' -f ($LogicalDisk_Item.Size / 1GB))
                    }
                }
            }
            catch {
                Write-Warning "[$($CimSession_Item.ComputerName)] Disk Check: $($PSItem.Exception.Message)"
            }
        }
    }
    end {
        if ($ComputerName -and $CimSession) { Remove-CimSession -CimSession $CimSession -ErrorAction Ignore }
    }
} #function