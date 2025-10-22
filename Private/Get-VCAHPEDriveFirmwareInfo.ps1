#Harold.Kammermeyer@vca.com
#Get drive firmware details from from iLO.
#Requires -Modules ImportExcel, VMware.VimAutomation.Core

. "$PSScriptRoot\Get-FirmwareVersion.ps1"

function Get-VCAHPEDriveFirmwareInfo {
    [CmdletBinding()]
    param(
        [parameter(
            Mandatory,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [alias('ComputerName', 'Name', 'CN')]
        [string[]]$ServerIlo,
        [parameter(Mandatory)]
        [pscredential]$Credential,
        [string]$FirmwareDBPath = "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\Firmware Update\HPServerFirmware.xlsx",
        [pscustomobject[]]$FirmwareDB,
        [VMware.Vim.HostSystem[]]$VIObject
    )

    begin {
        if (-not $PSBoundParameters.ContainsKey('FirmwareDB')) { $FirmwareDB = Import-Excel -Path $FirmwareDBPath }
        if (-not $PSBoundParameters.ContainsKey('Credential')) { $Credential = Get-Credential -Message 'Ilo Credentials:' }

        $ReachableIloList = Find-HPEiLO -Range $ServerIlo -WarningAction SilentlyContinue
        $IloConnection = Connect-HPEiLO -IP $ReachableIloList.Hostname -Credential $Credential -Disablecertificateauthentication -WarningAction SilentlyContinue

        if ($IloConnection -eq $null) {
            Write-Warning "Connection could not be established to any target iLO."
            Write-Host $IloConnection.Hostname | Format-List
            break
        }

        if ($IloConnection.count -ne $ServerIlo.count) {
            #List of IP's that could not be connected
            Write-Warning "Connection failed for below set of targets"
            foreach ($ServerIlo_Item in $ServerIlo) {
                if ($IloConnection.Hostname -notcontains $ServerIlo_Item) {
                    Write-Host $ServerIlo_Item | Format-List
                }
            }
        }

        $IloObj = Get-HPEiLOSmartArrayStorageController -Connection $IloConnection
    }
    process {
        $IloObj | ForEach-Object {
            $IloObj_Item = $PSItem
            $PSItem.Controllers.PhysicalDrives | ForEach-Object {
                $DriveFirmware = Get-FirmwareVersion -Model $PSItem.Model -FirmwareDB $FirmwareDB
                $Hostname = $IloObj_Item.Hostname -replace '-ilo', ''
                [pscustomobject]@{
                    Hostname           = $Hostname
                    IloHostname        = $IloObj_Item.Hostname
                    IloIP              = $IloObj_Item.IP
                    Location           = $PSItem.Location
                    LocationFormat     = $PSItem.LocationFormat
                    CapacityGB         = $PSItem.CapacityGB
                    InterfaceSpeedMbps = $PSItem.InterfaceSpeedMbps
                    RotationalSpeedRpm = $PSItem.RotationalSpeedRpm
                    InterfaceType      = $PSItem.InterfaceType
                    MediaType          = $PSItem.MediaType
                    Model              = $PSItem.Model
                    SerialNumber       = $PSItem.SerialNumber
                    FirmwareVersion    = $PSItem.FirmwareVersion
                    LatestFirmware     = $DriveFirmware.Firmware
                    State              = $PSItem.State
                    #ServerBootTime     = $VIObject.Where( { $_.Name -eq $Hostname } ).Runtime.BootTime
                    FirmwareStatus     = $(if ($PSItem.FirmwareVersion -eq $DriveFirmware.Firmware) { 'Up to date' } else { 'Check firmware' })
                    FirmwareFilename   = $DriveFirmware.FirmwareFilename
                    FirmwareFilePath   = $DriveFirmware.FirmwareFilepath
                    Error              = ''
                }
            }
        }
    } #process
    end {
        foreach ($ServerIlo_Item in $ServerIlo) {
            if ($IloConnection.Hostname -notcontains $ServerIlo_Item) {
                [pscustomobject]@{
                    Hostname           = $ServerIlo_Item -replace '-ilo', ''
                    IloHostname        = $ServerIlo_Item
                    IloIP              = ''
                    Location           = ''
                    LocationFormat     = ''
                    CapacityGB         = ''
                    InterfaceSpeedMbps = ''
                    RotationalSpeedRpm = ''
                    InterfaceType      = ''
                    MediaType          = ''
                    Model              = ''
                    SerialNumber       = ''
                    FirmwareVersion    = ''
                    LatestFirmware     = ''
                    State              = ''
                    #ServerBootTime     = $VIObject.Where( { $_.Name -eq $Hostname } ).Runtime.BootTime
                    FirmwareStatus     = ''
                    FirmwareFilename   = ''
                    FirmwareFilePath   = ''
                    Error              = 'Connection failed'
                }
            }
        } #foreach
    }
} #function