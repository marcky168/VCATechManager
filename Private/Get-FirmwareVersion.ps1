#Harold.Kammermeyer@vca.com
#Get Firmware version details from local excel spreadsheet.
#Requires -Modules ImportExcel

function Get-FirmwareVersion {
    [CmdletBinding()]
    param(
        [parameter(
            Position = 0)]
        [string[]]$Model,
        [string]$FirmwareDBPath = "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\Firmware Update\HPServerFirmware.xlsx",
        [pscustomobject[]]$FirmwareDB
    )
    begin {
        if (-not $PSBoundParameters.ContainsKey('FirmwareDB')) { $FirmwareDB = Import-Excel -Path $FirmwareDBPath }
    }
    process {
        foreach ($Model_Item in $Model) {
            $FirmwareDB.Where( { $PSItem.Model -eq $Model_Item } ) | Select-Object -Property Model, @{n = 'Firmware'; e = { $_.'LatestFWVer' } }, FirmwareFilename, FirmwareFilepath
        }
    } #process
    end {
        #intentionally left blank
    }
} #function