#Harold.Kammermeyer@vca.com
#Requires -Modules ActiveDirectory

function Get-OldUserProfiles {
    [cmdletbinding()]
    param(
        [alias('Name', 'CN')]
        [string[]]$ComputerName,
        [int]$TechRetentionDays = 30,
        [PSCredential]$Credential,
        [System.Management.Automation.Runspaces.PSSession[]]$Session
    )
    begin {
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
            'vcaantech.com/Corp Users/Medical Operations/*'
            'vcaantech.com/Corp Users/IT Systems Architecture/*'
            'vcaantech.com/Corp Users/IT Datacenter/*'
            #>
        )
    }
    process {
        $UserProfiles = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-CimInstance -ClassName Win32_UserProfile -Filter "Loaded=False And Special=False" | Where-Object SID -Like "S-1-5-21*" |
            Select-Object -Property LastUseTime, Loaded, LocalPath, SID, Special, @{n = 'LastWriteTime'; e = { (Get-Item -Path "$($_.LocalPath)\AppData\Local" -Force -ErrorAction SilentlyContinue).LastWriteTime } }
        } -Credential $Credential

        $ProfilesResolved = $UserProfiles | foreach-object {
            Clear-Variable -Name UserSID, UserADDisk -ErrorAction Ignore
            $UserSID = $_.SID

            try {
                $UserADDisk = Get-ADUser -Filter "SID -like '$($_.SID)'" -Properties CanonicalName -ErrorAction Stop
                [pscustomobject]@{
                    PSComputerName = $_.PSComputerName
                    Name           = if ($UserADDisk.Name) { $UserADDisk.Name } else { 'Could not resolve SID' }
                    SamAccountName = $UserADDisk.SamAccountName
                    LocalPath      = $_.LocalPath
                    LastUseTime    = $_.LastUseTime
                    LastWriteTime  = $_.LastWriteTime
                    Loaded         = $_.Loaded
                    SID            = $_.SID
                    Special        = $_.Special
                    CanonicalName  = $UserADDisk.CanonicalName
                }
            }
            catch {
                Write-Warning "$UserSID`: $($PSItem.Exception.Message)"
            }
        }

        # Profiles with unknown SID older than 60 days
        $ProfilesResolved | Where-Object { $_.Name -EQ 'Could not resolve SID' -and $_.LastWriteTime -lt (Get-Date).AddDays(-15) }
        # Profiles with disabled accounts
        $ProfilesResolved | Where-Object CanonicalName -like 'vcaantech.com/User Disabled Accounts/*'
        # Tech Profiles over $TechRetentionDays days old (defaulted to 30 days)
        $OldTechProfiles = $ProfilesResolved | Where-Object LastWriteTime -lt (Get-Date).AddDays(-$TechRetentionDays) | ForEach-Object {
            $ProfilesResolved_Item = $_
            $LdapPathAccountDeletion | ForEach-Object {
                if ($ProfilesResolved_Item.CanonicalName -like $_) {
                    $ProfilesResolved_Item
                }
            }
        }
        $OldTechProfiles
    }
    end {
        # intentionally left blank
    }
}