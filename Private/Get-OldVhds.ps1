#Harold.Kammermeyer@vca.com
#Requires -Modules ActiveDirectory

Function Get-OldVhds {
    [cmdletbinding()]
    param(
        [alias('Name', 'CN')]
        [string[]]$ComputerName,
        [string]$Path = 'H:\Users',
        [int]$TechRetentionDays = 30,
        [System.Management.Automation.Runspaces.PSSession[]]$Session
    )
    begin {
        $VhdFilesParam = @{}
        if ($Session) { $VhdFilesParam.Session = $Session }
        if ($ComputerName) { $VhdFilesParam.ComputerName = $ComputerName }

        $LdapPathAccountDeletion = @(
            'vcaantech.com/Admins/*'
            'vcaantech.com/Contractor Accounts/*'
            'vcaantech.com/Corp Users/*'
            'vcaantech.com/CAPNA/Corp Users/*'
            <#
            'vcaantech.com/Contractor Accounts/Corporate Contractors/*'
            'vcaantech.com/Contractor Accounts/Sparky Consultants/*'
            'vcaantech.com/Corp Users/Engineering/*'
            'vcaantech.com/Corp Users/Operations/*'
            'vcaantech.com/Corp Users/Service Desk/*'
            'vcaantech.com/Corp Users/Training/*'
            'vcaantech.com/Corp Users/IT Training/*'
            'vcaantech.com/Corp Users/Revenue Accounting/*'
            'vcaantech.com/Corp Users/Pricing/*'
            'vcaantech.com/Corp Users/Marketing/*'
            'vcaantech.com/Corp Users/Digital Client Experience/*'
            'vcaantech.com/Corp Users/IT EDO/*'
            #>
        )
    }
    process {
        try {
            # Get VHD file info, exclude template
            $VhdFiles = Invoke-Command @VhdFilesParam -ScriptBlock {
                (Get-ChildItem -Path $using:Path).Where( { ($_.Name -like "*.vhdx") -and ($_.Name -notlike "*template*") }) |
                Select-Object -Property Name, FullName, LastWriteTime, Length
            }
            # Resolve SID from ActiveDirectory
            $VhdResolved = $VhdFiles | ForEach-Object {
                Clear-Variable -Name UserSID, UserADDisk -ErrorAction Ignore
                $UserSID = [regex]::Match($PSItem.Name, '(?i)S-\d-\d+-(\d+-){1,14}\d+').Value

                try {
                    $UserADDisk = Get-ADUser -Filter "SID -like '$UserSID'" -Properties CanonicalName -ErrorAction Stop
                    [PSCustomObject]@{
                        Computer       = $PSItem.PSComputerName
                        Name           = if ($UserADDisk.Name) { $UserADDisk.Name } else { 'Could not resolve SID' }
                        SamAccountName = $UserADDisk.SamAccountName
                        SID            = $UserSID
                        FullName       = $PSItem.FullName
                        LastWriteTime  = $PSItem.LastWriteTime
                        Length         = $PSItem.Length
                        CanonicalName  = $UserADDisk.CanonicalName
                    }
                }
                catch {
                    Write-Warning "$UserSID`: $($PSItem.Exception.Message)"
                }
            }
            # Profiles with unknown SID older than 60 days
            $VhdResolved | Where-Object { $_.Name -EQ 'Could not resolve SID' -and $_.LastWriteTime -lt (Get-Date).AddDays(-15) }

            # Vhds with disabled accounts
            $VhdResolved | Where-Object CanonicalName -like 'vcaantech.com/User Disabled Accounts/*'
            # Tech vhds over $TechRetentionDays days old (defaulted to 30 days)
            $OldTechVhds = $VhdResolved | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$TechRetentionDays) | ForEach-Object {
                $VhdResolved_Item = $_
                $LdapPathAccountDeletion | ForEach-Object {
                    if ($VhdResolved_Item.CanonicalName -like $_) {
                        $VhdResolved_Item
                    }
                }
            }
            $OldTechVhds
        }
        catch {
            throw $_
        }
    }
    end {
        # intentionally left blank
    }
}