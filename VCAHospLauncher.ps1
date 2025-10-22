#Harold.Kammermeyer@vca.com
#Requires -Version 5

# Check if script is being launched from network share; if so add vcaantech.com to trusted domain sites if not already
if (([System.Uri]$PSCommandPath).IsUnc) {
    $TrustedSitesPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\vcaantech.com\'

    if (-not (Get-Item -Path $TrustedSitesPath -ErrorAction Ignore)) {
        New-Item -Path $TrustedSitesPath | Out-Null
        New-ItemProperty -Path $TrustedSitesPath -Name file -Value 2 -Type DWORD | Out-Null
    }
}

Import-Module -Name ActiveDirectory -WarningAction Ignore -Verbose
#Import-Module -Name DnsServer -WarningAction Ignore
#Import-Module -Name HPEiLOCmdlets -WarningAction Ignore

Import-Module -Name "$PSScriptRoot\Private\lib\ColorHost" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\ImportExcel" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\CredentialManager" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\Posh-SSH" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\Autoload" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\Reflection" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\Wasp" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\PnP.PowerShell" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\HPEiLOCmdlets" -MinimumVersion 3.3.0.0 -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\HPEBIOSCmdlets" -MinimumVersion 3.0.0.0 -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\AdmPwd.PS" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\Proxx.SNMP" -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\PSMenu"  -Verbose
Import-Module -Name "$PSScriptRoot\Private\lib\QuserObject" -Verbose

if (-not (Get-Module -Name VMware.VimAutomation.Core -ListAvailable)) {
    Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Sdk" -Verbose
    Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Common" -Verbose
    Import-Module -Name "$PSScriptRoot\Private\lib\VMware.Vim" -Verbose
    Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Cis.Core" -Verbose
    Import-Module -Name "$PSScriptRoot\Private\lib\VMware.VimAutomation.Core" -Verbose
}
if ((Get-Module -Name VMware.VimAutomation.Core)) {
    if ((Get-PowerCLIConfiguration -Scope User).ParticipateInCEIP -ne $false) {
        Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false | Out-Null
    }
    if ((Get-PowerCLIConfiguration -Scope User).InvalidCertificateAction -eq 'Unset') {
        Set-PowerCLIConfiguration -Scope User -InvalidCertificateAction 'Warn' -Confirm:$false | Out-Null
    }
}

# Adding certificate exception
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

. "$PSScriptRoot\Private\Convert-VcaAU.ps1"
. "$PSScriptRoot\Private\Get-DriveStatus.ps1"
. "$PSScriptRoot\Private\Get-VCAHeadCount.ps1"
. "$PSScriptRoot\Private\Get-MemoryUsage.ps1"
. "$PSScriptRoot\Private\Get-DiskUsage.ps1"
. "$PSScriptRoot\Private\Get-UPSStatus.ps1"
. "$PSScriptRoot\Private\whatusers.ps1"
. "$PSScriptRoot\Private\whatdisk.ps1"
. "$PSScriptRoot\Private\Remove-BakRegistry.ps1"
. "$PSScriptRoot\Private\New-ServiceNowIncident.ps1"
. "$PSScriptRoot\Private\New-ServiceNowGUI.ps1"
. "$PSScriptRoot\Private\Get-RdsConnectionConfig.ps1"
. "$PSScriptRoot\Private\Get-OldVhds.ps1"
. "$PSScriptRoot\Private\Get-OldUserProfiles.ps1"
. "$PSScriptRoot\Private\Get-VCAHPEDriveFirmwareInfo.ps1"
. "$PSScriptRoot\Private\Get-FirmwareVersion.ps1"
. "$PSScriptRoot\Private\Get-ADUserLockouts.ps1"
. "$PSScriptRoot\Private\Copy-ToPSSession.ps1"
. "$PSScriptRoot\Private\Get-QuserStateParallel.ps1"
. "$PSScriptRoot\Private\Test-ConnectionAsync.ps1"
. "$PSScriptRoot\Private\Get-WindowsNetwork.ps1"

function VCAOpsPortal {
    [CmdletBinding()]
    param(
        [string]$ComputerName,
        [switch]$NoBannerMessage,
        [switch]$NoMenu,
        [switch]$NoVersionCheck
    )

    [decimal]$Version = 251016
    $ScriptWriteTime = (Get-ChildItem -Path $PSCommandPath).LastWriteTime

    # Clear console if no error occurred
    Write-Host $Error.count
    if ($Error.Count -le 50) { [System.Console]::Clear() }

    # Set title
    $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - $PSCommandPath (Written:$ScriptWriteTime)"
    if (-not $NoBannerMessage.IsPresent) {
        Write-Host "`r`nVCA Ops Portal v.$Version`r`n" -ForegroundColor Yellow
    }
    # Version check
    if (-not $NoVersionCheck.IsPresent) { Get-VcaOpsPortalVersion -Version $Version -OnlyNew }

    # Load hospital master to memory
    if (-not $script:HospitalMaster) {
        $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
    }
    # Load clusters list to memory
    if (-not $script:ClusterSites) { $script:ClusterSites = Import-Csv -Path "$PSScriptRoot\Private\csv\ClusterSites.csv" }
    # Load credentials
    $script:EsxiCredential = Get-StoredCredential -Target vcahospesxi
    $script:IloCredential = Get-StoredCredential -Target vcahospilo
    $script:ADCredential = Get-StoredCredential -Target vcadomaincreds
    $script:SNOWAPICredential = Get-StoredCredential -Target vcasnowapi
    $script:EmailCredential = Get-StoredCredential -Target vcaemailcreds
    # Load Ops Full Menu Description Csv
    $script:PortalMenuCsv = Import-Csv -Path "$PSScriptRoot\Private\lib\Menu.csv"
    # Load UPS SNMP OID Description Details
    $HPUpsSnmp = Import-Csv -Path "$PSScriptRoot\Private\lib\HP UPS SNMP.csv"

    while (-not $ComputerName) {
        # Reset powershell window title
        $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - $PSCommandPath (Written:$ScriptWriteTime)"

        Clear-Variable -Name MenuAction -ErrorAction Ignore

        # Start AU User Prompt loop
        if (-not $ComputerNameSelection) {
            # ComputerName prompt
            Write-Host 'Enter Hospital AU <Number>, <Hostname>, <IP>, (M)aster, (C)ircuit, (robo), (rcan|rc|r)andom: ' -ForegroundColor Cyan -NoNewline
            $ComputerName = (Read-Host).Trim()
        }
        else {
            # History Selection
            $ComputerName = Convert-VcaAU -AU $ComputerNameSelection -Strip
            if (-not $ComputerName) { $ComputerName = $ComputerNameSelection }
            Clear-Variable -Name ComputerNameSelection
        }

        if ($ComputerName -in 'master','m') { #m
            # Hospital Master Site Selection
            if (-not $HospitalMaster) {
                $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
            }

            if ($HospitalMaster) {
                Clear-Variable -Name MasterSelect -ErrorAction Ignore

                $HospitalMaster | Out-Gridview -Title "#m Hospital Master - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable MasterSelect | Out-Null
                if ($MasterSelect) { $ComputerName = $MasterSelect.'Hospital Number'.Trim() }

                if ($ComputerName -in 'master','m') { Clear-Variable -Name ComputerName }
            }
        }
        elseif ($ComputerName -in 'circuit','c') {
            # Circuit Site Selection
            if (-not $HospitalCircuits) {
                $script:HospitalCircuits = Import-Csv -Path "$PSScriptRoot\Private\csv\All-Hospital-Circuits.csv"
            }
            if ($HospitalCircuits) {
                Clear-Variable -Name CircuitSelect -ErrorAction Ignore

                $HospitalCircuits | Out-Gridview -Title "#c Hospital Master - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable CircuitSelect | Out-Null
                if ($CircuitSelect) { $ComputerName = $CircuitSelect.AU.Trim() }

                if ($ComputerName -in 'circuit','c') { Clear-Variable -Name ComputerName }
            }
        }
        elseif ($Computername -eq '37') {
            # WW API Health Report
            if (-not $SNOWAPICredential) { $SNOWAPICredential = Get-StoredCredential -Target vcasnowapi }
            Get-WWAPIHealth -Credential $SNOWAPICredential -HospitalMaster $HospitalMaster
            Clear-Variable -Name ComputerName -ErrorAction Ignore
        }
        elseif ($Computername -eq 'robo') {
            Invoke-VcaOpsPortalUpdate
            Clear-Variable -Name ComputerName -ErrorAction Ignore
        }
        elseif ($Computername -eq 'r') {
            # Set to random VCA site excluding canada and '0'
            $ComputerName = Get-Random -InputObject $($HospitalMaster.'Hospital Number'.where({$_ -notmatch '^8[0-9]{4}$|^h21[0-9]{4}$|^0$'}))
        }
        elseif ($Computername -eq 'rc') {
            $ClusterSitesFiltered = $ClusterSites.Where({$_.AU -notmatch '^8[0-9]{4}$|^h21[0-9]{4}$|^0$' -and $_.Cluster -notmatch "Cloud|Migration"}) | Select-Object -ExpandProperty AU -Unique
            $ComputerName = Get-Random -InputObject $(($ClusterSitesFiltered))
        }
        elseif ($Computername -eq 'rcan') {
            $ComputerName = Get-Random -InputObject $($HospitalMaster.'Hospital Number'.where({$_ -match '^8[0-9]{4}$|^h21[0-9]{4}$'}))
        }

        # Hostname, AU #, or IP match
        if ($ComputerName -match '^[a-zA-Z0-9-.]+$|^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$' -and $ComputerName -notin '0', '999', 'h') {
            Clear-Variable -Name ComputerDNS, Cluster, ClusterMatch, SiteAUNumber, VMCSDDC -ErrorAction Ignore

            # AU format match
            if ($ComputerName -match '^\d{3,6}$') {
                $SiteAUNumber = Convert-VcaAU -AU $ComputerName -Prefix 'AU' -Suffix '' -NoLeadingZeros
                $ClusterMatch = $ClusterSites.Where( { $PSItem.AU -eq $ComputerName } )

                if ($ComputerName -notin $ClusterSites.AU) {
                    $ComputerName = Convert-VcaAU -AU $ComputerName -EsxiHost
                }
                elseif (@($ClusterMatch).count -ge 2) {
                    #Cluster
                    $Cluster = $ClusterMatch.Name | Sort-Object
                }
                elseif (@($ClusterMatch).count -eq 1) {
                    $ComputerName = Convert-VcaAU -AU $ComputerName -EsxiHost
                    $VMCSDDC = $true
                }
            }
            # Hostname match
            if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$') {
                # Resolve Hostname to IP
                if (-not $Cluster) {
                    try {
                        # Autocorrect case on vmIncom servers otherwise Action 6 fails to go to directory due to linux case senstivity.
                        $ComputerName = $ComputerName.Replace('vmincom', 'vmIncom')
                        $ComputerDNS = (Resolve-DnsName -Name $ComputerName -ErrorAction Stop).IPAddress
                        Write-Host "`r`n Server: $ComputerName ($ComputerDNS)$(' '*70)`r`n" -ForegroundColor Black -BackgroundColor Yellow
                    }
                    catch {
                        $ComputerDNS = $_.Exception.Message
                        Write-Host "`r`n Server: $ComputerName ($ComputerDNS)$(' '*45)`r`n" -ForegroundColor White -BackgroundColor DarkRed
                    }
                }
                else {
                    # Cluster
                    Clear-Variable -Name ClusterHostCount, ClusterHost -ErrorAction Ignore
                    $ComputerDNS = $Cluster | Resolve-DnsName -ErrorAction SilentlyContinue
                    $ClusterSite = ($ClusterSites.Where( { $PSItem.AU -eq $ComputerName })).Cluster[0]
                    $ClusterHostCount = 1
                    Write-Host "`r`n 0. Change site"
                    $ClusterVMHost = $ComputerDNS | foreach-object {
                        Write-Output " $ClusterHostCount. Server: $ClusterSite - $($PSItem.Name) ($($PSItem.IPAddress)) "
                        $ClusterHostCount += 1
                    }
                    Write-Output $ClusterVMHost | Out-String

                    $Subnet = (([ipaddress]$ComputerDNS[0].IPAddress).GetAddressBytes()[0..2] -join '.') + '.'
                    while ((-not $ClusterHost) -or ($ClusterHost -notmatch "^[0-$($Cluster.Count)]$")) {
                        Write-Host "[$ClusterSite$(if ($Subnet) { " ($Subnet)" })] Select VMHost: " -ForegroundColor Cyan -NoNewline
                        $ClusterHost = (Read-Host).Trim()
                    }
                    switch -regex ($ClusterHost) {
                        '^0$' {
                            $MenuAction = '0'
                            Write-Host ''
                            Clear-Variable -Name ComputerName
                        }
                        "^[1-$($Cluster.Count)]$" {
                            $ComputerName = ($ClusterVMHost.Where( { $PSItem -like " $ClusterHost. Server:*" } ) -split ' - ')[1].Split('.')[0]
                            try {
                                $ComputerDNS = ($ComputerName | Resolve-DnsName -ErrorAction Stop).IPAddress
                                Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*70)`n" -ForegroundColor Black -BackgroundColor Yellow
                            }
                            catch {
                                $ComputerDNS = $_.Exception.Message
                                Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*45)`n" -ForegroundColor White -BackgroundColor DarkRed
                            }
                        }
                    } #switch -regex ($ClusterHost)
                } #else
            } #if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$')
            else {
                # Matches IP format
                Write-Host "`r`n Server: $ComputerName$(' '*90)`r`n" -ForegroundColor Black -BackgroundColor Yellow
            }

            # Set title with hostname/timezone
            if ($HospitalMaster -and $ComputerName) {
                $HospitalInfo = $HospitalMaster.Where( {
                        $PSItem.'Hospital Number' -eq "$(Convert-VcaAU -AU $ComputerName -Strip)" } )
            }
            $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - [$ComputerName] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available '}) - $PSCommandPath (Written:$ScriptWriteTime)"

            Clear-Variable -Name SiteADComputers, NetServicesObj, NetServices, SingleHost -ErrorAction Ignore
            # Set $NetServices with all '*-NS*' computer object matches in AD
            if (($ComputerName -match '-vm|hmtprod-') -and ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$')) {
                $SiteADComputers = Get-VcaADComputers -AU $ComputerName
                $NetServicesObj = $SiteADComputers | Where-Object Name -match "-ns|-ra"
                $NetServices = $NetServicesObj | Select-Object -ExpandProperty Name
            }
            else {
                $NetServices = $ComputerName
                $SingleHost = $true
            }

            Clear-Variable -Name NSIPAddress -ErrorAction Ignore
            # Set title with VMC subnet info
            if (-not $Cluster -and -not $SingleHost) {
                try {
                    # Resolve $NetServices Hostnames to an IP addreess
                    $NSIPAddress = [System.Net.Dns]::GetHostAddresses(@($NetServices)[0]).IPAddressToString

                    if (-not $SiteAUNumber -and $NSIPAddress -match "^10.242.|^10.225.") {
                        $SiteAUNumber = Convert-VcaAU -AU $ComputerName -Prefix 'AU' -Suffix '' -NoLeadingZeros
                    }
                    if ($NSIPAddress -like "10.242.*") {
                        if (@($NetServicesObj).count -eq 1) {
                            $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - [$SiteAUNumber - $($NetServicesObj.Subnet)] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available '}) - $PSCommandPath (Written:$ScriptWriteTime)"
                        }
                        elseif (@($NetServicesObj).count -ge 2) {
                            $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - [$SiteAUNumber - $($NetServicesObj[0].Subnet)] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available '}) - $PSCommandPath (Written:$ScriptWriteTime)"
                        }
                    }
                    elseif ($NSIPAddress -like "10.225.*") {
                        $Host.UI.RawUI.WindowTitle = "VCA Ops Portal v.$Version - [$SiteAUNumber - Mesa] - $(if ($HospitalInfo.'Time Zone') { $HospitalInfo.'Time Zone'} else { 'Timezone Not Available '}) - $PSCommandPath (Written:$ScriptWriteTime)"
                    }
                    <#
                    if ($NSIPAddress -match "^10.242.|^10.225.") {
#                        Get-VcaADComputers -AU $ComputerName | Format-Table -AutoSize | Out-String
                        Get-VcaADComputers -AU $ComputerName | Out-TableString -Wrap
                    }
                    #>
                }
                catch {
                    # intentionally left blank
                }
            }
            #Menu
            if (-not $NoMenu.IsPresent) {
if (-not $UserGroups) { $UserGroups = (whoami /groups) }
$TelecomTeam = @(
    'Patrick.Perkins'
    'Victor.Marquez'
    'Gilbert.Jaime'
    'Yvette.Muro'
    'Josh.Bush'
    'Phil.Rhymer'
)
$WOOFWareTeam = @(
    'Adel.Buckley'
    'Melanie.West'
    'Tammie.Wojcik'
    'Marc.Yap'
)
    $OpsMenu = @"
`r Select action:
         0 - $[37mchange host$[0m                                     21 - $[37m net: solarwinds orion$[0m
     1,L,s - esxi: drive report (notepad)                  22,w - $[37m ref: site pictures$[0m
         2 - esxi: drive report (screen only)                23 -  net: display dns CNAME and A records
         3 - esxi: enable ssh                            24,abc -  net: dhcp mgmt; a:Options; b:Leases; c:Reservations
         4 - esxi: disable ssh                               25 - $[37m win: restart windows services$[0m
         5 - esxi: launch putty (ssh)                        26 - $[93m ref: reference portals/links$[0m
         6 - esxi: launch scp                             27,dp -  ilo: query ilo status/disks/powermode
 7,77[L|s] - $[37m win: quser/launch vnc/launch RDP Shadow$[0m      28,n - esxi: query esxi host/network
         8 - $[37m  ad: site users & headcount$[0m                    29 - esxi: vmware remote console
         9 -  ilo: launch ilo web                            30 - $[37m win: check critical services (windows)$[0m
        10 - esxi: launch esxi webui                       31,w - $[37m ups: query server ups$[0m
        11 - $[37m win: memory/uptime/disk/cpu usage$[0m              32 -  ilo: ilo remote console
        12 -  win: expand windows drive                    33,t -  ilo: server uid status/toggle uid
      13,p - $[37msite: site ping status/gateway quality$[0m          34 -  ilo: view iml
        14 - $[37m ref: hospital contact/hours$[0m           35,i,o,e,oe - $[37m ref: launch service now/ops/extended$[0m
        15 - $[37m ref: all-hospital-circuits (from web)$[0m        36,a -  ilo: reset ilo/aux cycle
        16 - $[37m ref: launch hospital website/info$[0m          37,c,i - $[93mreport: WOOFware api health report(i)ndividual$[0m
        17 -  ilo: find-hpeilo                               38 - $[37m win: query installed applications$[0m
        18 - $[37msite: ip scanner$[0m                              40,s -  win: PortQuery (Outbound Connectivity)
      19,? - $[37m win: rdc (clusters - /admin)$[0m                   41 -  win: find temp folders
        20 - $[37m  ad: hosp. servers in active directory$[0m         42 -  win: stop appreadiness

                                                           98,L - $[37mescalation numbers/Escalation document$[0m
                                                         99,L,r - $[37mmanage credentials/LAPS/refresh$[0m
                                                            999 - $[37mnew session$[0m
                                                            000 - $[37mchangelog$[0m
                                                           ??,! - $[37mshow all available actions/gui$[0m`r`n
"@
if ($UserGroups -match 'SN_Operations|SN_Engineering-Hospital|Domain Admins') {
    $Menu = $OpsMenu
}
elseif (($UserGroups -match 'SN_Telecom') -or ($env:USERNAME -in $TelecomTeam)) {
    $Menu = @"
`r Select action:
      0 - $[37mchange host$[0m                                     60 - $[93mOpen Telecom's Template Foler$[0m
  7,77s - $[37mquser/launch vnc$[0m                             60pbx - $[93mDisplay PBX (NEC) & UPS information$[0m
      8 - $[37msite users & headcount$[0m                       60sow - $[93mLaunch Telecom's SOW Template (Word Document)$[0m
     13 - $[37msite ping status$[0m                              60it - $[93mLaunch Telecom's AT Installation Info (Excel Document)$[0m
     14 - $[37mhospital master$[0m                              60con - $[93mLaunch Telecom's Contacts (List)$[0m   
     15 - $[37mall-hospital-circuits (latest from web)$[0m     60carr - $[93mLaunch Telecom's Line Carriers (Excel Document)$[0m
     16 - $[37mlaunch hospital website$[0m                     60toll - $[93mLaunch Telecom's Toll Free Inventory (Excel Document)$[0m
     18 - $[37mip scanner$[0m                                      61 - $[93mLaunch 8x8 'Live' Hospitals (SmartSheet)$[0m
     20 - $[37mhosp. servers in active directory$[0m
     21 - $[37msolarwinds orion$[0m
     22 - $[37msite pictures$[0m
     23 - display dns A records
 24,abc - net: dhcp mgmt; a:Options; b:Leases; c:Reservations    98,L - escalation numbers/Escalation document
     26 - $[37mlaunch web portals$[0m                          99,L,r - manage credentials/LAPS/refresh
     31 - $[37mcheck server ups$[0m                               999 - $[37mnew session$[0m
     35 - $[37mlaunch service now$[0m                             000 - changelog`r`n
"@
}
elseif (($UserGroups -match 'SN_WOOFWare Support') -or ($env:USERNAME -in $WOOFWareTeam)) {
    $Menu = @"
`r Select action:
    0 - change host                                               21 - net: solarwinds orion
 ping -  win: Send ping from selected VM for key hosts          22,w - ref: site pictures
   up -  win: Get last reboot time for selected Guest VMs     24,abc - net: dhcp mgmt; a:Options; b:Leases; c:Reservations
 7,77 -  win: quser/launch vnc                                    25 - win: restart windows services
  77l -  win: Log user off                                        26 - $[93mref: reference portals/links$[0m
  77s -  win: Shadow user terminal session                        30 - win: check critical services (windows)
    8 -   ad: site users & headcount                            31,w - ups: query server ups
   8t - snow: Select impacted user for ServiceNow incident    35,i,e - ref: launch service now/extended
   11 -  win: memory/uptime/disk/cpu usage                        38 - win: query installed applications
  11m -  win: Display memory usage by user                      40,s - win: PortQuery (Outbound Connectivity)
  11h -  win: Display memory usage for "Heavy Hitters"            41 - win: find temp folders
  11p -  win: Display Process memory usage by user               41d - win: Display disks mounted on NS Clusters
  11v -  win: Display WOOFware memory usage by user               42 - win: stop appreadiness
11time-  win: Retrieve NS local time                              43 - win: find .bak profiles
 13,p - site: site ping status/gateway quality                   43f - win: Purge .bak files from registry
   14 -  ref: hospital contact/hours                              50 - win: Clear print spooler
  14t - snow: Generate ServiceNow incident                        51 - win: Query installed printers/Launch printer web management
   15 -  ref: all-hospital-circuits (from web)                   70a - win: Check latest 100 events in Application event logs
   16 -  ref: launch hospital website/info                       70s - win: Check latest 100 events in System event logs
   18 - site: ip scanner                                          80 - ww: Clear GMR Cache
 19,? -  win: rdc (clusters - /admin)                             81 - Launch WOOFware Reports Website
   20 -   ad: hosp. servers in active directory                   82 - Launch Fuse Website
                                                                  83 - win: Restart Sparky Services

                                                                98,L - $[37mescalation numbers/Escalation document$[0m
                                                              99,L,r - $[37mmanage credentials/LAPS/refresh$[0m
                                                                 999 - $[37mnew session$[0m
                                                                 000 - $[37mchangelog$[0m
                                                                ??,! - $[37mshow all available actions/gui$[0m`r`n
"@
}
else {
    # Default menu
    $Menu = @"
`r Select action:
    0 - change host                                               21 - net: solarwinds orion
 ping -  win: Send ping from selected VM for key hosts          22,w - ref: site pictures
   up -  win: Get last reboot time for selected Guest VMs     24,abc - net: dhcp mgmt; a:Options; b:Leases; c:Reservations
 7,77 -  win: quser/launch vnc                                    25 - win: restart windows services
  77l -  win: Log user off                                        26 - $[93mref: reference portals/links$[0m
  77s -  win: Shadow user terminal session                        30 - win: check critical services (windows)
    8 -   ad: site users & headcount                            31,w - ups: query server ups
   8t - snow: Select impacted user for ServiceNow incident    35,i,e - ref: launch service now/extended
   11 -  win: memory/uptime/disk/cpu usage                        38 - win: query installed applications
  11m -  win: Display memory usage by user                      40,s - win: PortQuery (Outbound Connectivity)
  11h -  win: Display memory usage for "Heavy Hitters"            41 - win: find temp folders
  11p -  win: Display Process memory usage by user               41d - win: Display disks mounted on NS Clusters
  11v -  win: Display WOOFware memory usage by user               42 - win: stop appreadiness
11time-  win: Retrieve NS local time                              43 - win: find .bak profiles
 13,p - site: site ping status/gateway quality                   43f - win: Purge .bak files from registry
   14 -  ref: hospital contact/hours                              50 - win: Clear print spooler
  14t - snow: Generate ServiceNow incident                        51 - win: Query installed printers/Launch printer web management
   15 -  ref: all-hospital-circuits (from web)                   70a - win: Check latest 100 events in Application event logs
   16 -  ref: launch hospital website/info                       70s - win: Check latest 100 events in System event logs
   18 - site: ip scanner
 19,? -  win: rdc (clusters - /admin)
   20 -   ad: hosp. servers in active directory

                                                                98,L - $[37mescalation numbers/Escalation document$[0m
                                                              99,L,r - $[37mmanage credentials/LAPS/refresh$[0m
                                                                 999 - $[37mnew session$[0m
                                                                 000 - $[37mchangelog$[0m
                                                                ??,! - $[37mshow all available actions/gui$[0m`r`n
"@
}
                if ($MenuAction -ne '0') {
                    Write-ColorHost $Menu -ForegroundColor Cyan

                    Clear-Variable -Name VMCSDDCCluster -ErrorAction Ignore
                    # Display AD computers and subnet location for non on-prem sites
                    if (-not $VMCSDDC -and -not $Cluster) {
                        if ($NSIPAddress -like "10.242.*") {
                            Write-Host "* $(@($NetServices)[0]) ($NSIPAddress) on VMWare Cloud (SDDC) subnet`r`n" -ForegroundColor Magenta
                        }
                        elseif ($NSIPAddress -like "10.225.*") {
                            Write-Host "* $(@($NetServices)[0]) ($NSIPAddress) on Mesa Datacenter Subnet`r`n" -ForegroundColor Cyan
                        }
                        if ($NSIPAddress -match "^10\.242\.|^10\.225\.") {
                            $SiteADComputers | Out-TableString -NoNewLine -Wrap
                            if ($SiteADComputers.Name -match '-fs') {
                                $VMCSDDCCluster = $true
                            }
                        }
                    }
                    elseif ($VMCSDDC) {
                        Write-Host "* Site has been migrated or in process to be migrated to VMWare Cloud (SDDC)`r`n" -ForegroundColor Magenta
                        $SiteADComputers | Out-TableString -NoNewLine
                        if ($SiteADComputers.Name -match '-fs') {
                            $VMCSDDCCluster = $true
                        }
                    }
                }
            }
            while (-not $MenuAction) {
                # Action prompt
                Write-Host "[" -ForegroundColor Yellow -NoNewline
                Write-Host "$(Get-Date -Format "h:mm tt")" -ForegroundColor Gray -NoNewline
                Write-Host "][$ComputerName$(if ($ComputerDNS) { " ($ComputerDNS)" }) ?]: " -ForegroundColor Yellow -NoNewline
                $MenuAction = (Read-Host).Trim()
                if ($MenuAction -and ($PortalMenuCsv.Action | Where-Object { $_ -match [regex]::escape($MenuAction) })) { ( "v.$Version | $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" | Out-String) -replace '\r\n\r\n\r\n', '' }

                # Menu
                if ($MenuAction -match '^help$|^menu$|^\?$|^666$') {
                    if ($ComputerDNS -match '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') {
                        Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*70)`n" -ForegroundColor Black -BackgroundColor DarkYellow
                    }
                    else {
                        Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*45)`n" -ForegroundColor White -BackgroundColor DarkRed
                    }
                    Write-ColorHost $Menu -ForegroundColor Cyan
                    Clear-Variable -Name MenuAction
                }
                elseif ($MenuAction -match '^cls$|^clear$') {
                    Clear-Host
                    if ($ComputerDNS -match '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') {
                        Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*70)`n" -ForegroundColor Black -BackgroundColor Yellow
                    }
                    else {
                        Write-Host "`n Server: $ComputerName ($ComputerDNS)$(' '*45)`n" -ForegroundColor White -BackgroundColor DarkRed
                    }
                    Write-ColorHost $Menu -ForegroundColor Cyan
                    Clear-Variable -Name MenuAction
                }
                elseif ($MenuAction -match '^time$|^t$') {
                    (Get-Date | Out-String) -replace '\r\n\r\n\r\n', '' -replace '\r\n', ''
                    Clear-Variable -Name MenuAction
                }
                # Version check
                elseif ($MenuAction -match '^ver$|^version$' ) {
                    Write-Host "`r`nVCA Ops Portal v.$Version`r`n" -ForegroundColor Cyan
                    Get-VcaOpsPortalVersion -Version $Version -ShowWarning
                    Clear-Variable -Name MenuAction
                }
                else {
                    # Add ESXi Host to history
                    if (-not $SelectionHistory) { [System.Collections.ArrayList]$SelectionHistory = @() }
                    if ($SelectionHistory[-1].Name -ne $ComputerName) {
                        $SelectionHistory.Add(($ComputerName |
                        Select-Object -Property @{n='Timestamp';e={Get-Date}},
                                                @{n='Name';e={$PSItem}},
                                                @{n='IP';e={$ComputerDNS}})) | Out-Null
                    }
                    # Parse actions
                    switch ($MenuAction) {
                        '??' {
                            # #??
                            # Display all commands
                            $PortalMenuCsv | Out-TableString -Wrap
                        }
                        '??!' {
                            # #??!
                            Clear-Variable -Name PortalMenuSelection -ErrorAction Ignore
                            $PortalMenuSelection = $PortalMenuCsv | Out-GridView -Title "#??! Full Portal Tool Actions list - Select Action(s) to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -PassThru
                            if ($PortalMenuSelection) { $PortalMenuSelection | Out-TableString }
                        }
                        'd' {
                            # #d
                            # Runs a diagnostic bundle of commands
                            if (-not $ADCredential) { $ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            if (-not $IloCredential ) { $IloCredential = Get-StoredCredential -Target vcahospilo }
                            Clear-Variable -Name VCAStdTSNames -ErrorAction Ignore

                            $VCAStdTSNames = Get-VcaStdTSNames -ComputerName $ComputerName | Select-Object -ExpandProperty Name
                            Get-GuestResource -ComputerName (@($NetServices) + @($VCAStdTSNames)) -Credential $script:ADCredential
                            Get-Quser -ComputerName $Netservices
                            Get-UserMemory -NetServices $NetServices
                            Get-IloStatus -NetServices $NetServices -Cluster $Cluster -ComputerName $ComputerName -Credential $IloCredential
                        }
                        'e' {
                            # #e
                            Invoke-Expression 'cmd /c start powershell -NoExit -Command { . "\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Archive\VCA Ops Portal-181025\Reports\Invoke-PSHtml5.ps1" }'
                        }
                        'e2' {
                            # #e2
                            Invoke-Expression 'cmd /c start powershell -WindowStyle Hidden -Command {
                                [console]::beep(440,500)
                                [console]::beep(440,500)
                                [console]::beep(440,500)
                                [console]::beep(349,350)
                                [console]::beep(523,150)
                                [console]::beep(440,500)
                                [console]::beep(349,350)
                                [console]::beep(523,150)
                                [console]::beep(440,1000)
                                [console]::beep(659,500)
                                [console]::beep(659,500)
                                [console]::beep(659,500)
                                [console]::beep(698,350)
                                [console]::beep(523,150)
                                [console]::beep(415,500)
                                [console]::beep(349,350)
                                [console]::beep(523,150)
                                [console]::beep(440,1000)
                                }'
			}
        
                        'grdp' {
                            Clear-Variable  -name computers, service, brokerstart, brokerselection -ErrorAction Ignore

                            $computers = Get-ADComputer -Filter '((Name -like "h*-ns" -or Name -like "h*-fs*") -and Name -notlike "*-old") -and OperatingSystem -like "*Server*" -and Enabled -eq $true' -Properties Name | Select-Object -ExpandProperty Name | Sort-Object
                            $TitleTime = Get-Date -Format "yyyy-MM-dd HH:mm"
                            
                            $computers | Start-RSJob -Name BrokerJobs -Throttle 64 -ScriptBlock {
                                Invoke-Command -ComputerName $_ -ScriptBlock {
                                    Get-Service -Name tssdis -ErrorAction SilentlyContinue | Select-Object Status, Name, Pscomputername
                                }
                            } | Out-Null
                            $service = Get-RSJob -Name BrokerJobs | Wait-RSJob -ShowProgress -Timeout 600 | Receive-RSJob
                            Get-RSJob -Name BrokerJobs | Remove-RSJob -Force
                            
                            $brokerselection = $service | Select-Object -Property PSComputerName, Name, Status | Sort-Object -Property Status | Out-GridView -Title "#GRDP Global RDP Broker Check (Highlight row(s) that are Stopped and click OK ...to start service) - $Titletime - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -PassThru
                            
                            
                            if ($brokerselection) {
                                $brokerstart = $service | Where-Object { [string]$_.Status -eq 'Stopped' }
                                if ($brokerstart.PSComputerName -eq $null) { 
                                    Write-Host "`r`nThe selection(s) made didn't contain server(s) with a stopped service.  Aborting.`r`n" -ForegroundColor Red
                                } else {
                                    Invoke-Command -ComputerName $brokerstart.PSComputerName -ScriptBlock { $env:COMPUTERNAME; Start-Service -Name tssdis -Verbose -Confirm:$false }
                                }
                            }
                        }
                        'etsn' {
                            # #etsn
                            Write-Host "`r`n > Enter-PSSession -ComputerName $(@($NetServices)[0])`r`n" -ForegroundColor Cyan
                            Invoke-Expression -Command "cmd /c start powershell -NoExit -Command { Enter-PSSession -ComputerName $(@($NetServices)[0]) }"
                        }
                        'etsns' {
                            # #etsns
                            $SiteServers = Select-VcaSite -AU $ComputerName -Title "#etsns Select Server(s) to launch PSSession - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                            if ($SiteServers) {
                                $SiteServers.Name | foreach-object {
                                    Write-Host "`r`n > Enter-PSSession -ComputerName $_" -ForegroundColor Cyan
                                    Invoke-Expression -Command "cmd /c start powershell -NoExit -Command { Enter-PSSession -ComputerName $_ }"
                                }
                                Write-Host ''
                            }
                        }
                        'h' {
                            # #h
                            # Host history switcher
                            Clear-Variable -Name ComputerNameSelection, HostSwitchSelection -ErrorAction Ignore
                            $SelectionHistory | Sort-Object Timestamp -Descending | Out-GridView -Title "#h Select host to switch to - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable HostSwitchSelection | Out-Null
                            if ($HostSwitchSelection) {
                                $ComputerNameSelection = $HostSwitchSelection.Name

                                # Add ESXi Host to history
                                if (-not $SelectionHistory) { [System.Collections.ArrayList]$SelectionHistory = @() }
                                $SelectionHistory.Add(($ComputerName |
                                    Select-Object -Property @{n='Timestamp';e={Get-Date}},
                                                            @{n='Name';e={$PSItem}},
                                                            @{n='IP';e={$ComputerDNS}})) | Out-Null

                                Clear-Variable -Name ComputerName
                            }
                        }
                        'icm' {
                            # #icm
                            Clear-Variable -Name RemoteCommand -ErrorAction Ignore
                            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

                            $InputBoxTitle = "Remote command ($($Netservices))"
                            $InputBoxMessage = 'Enter a command:'
                            $InputBoxDefault = 'ping internic.net'

                            $RemoteCommand = [Microsoft.VisualBasic.Interaction]::InputBox($InputBoxMessage, $InputBoxTitle, $InputBoxDefault)

                            if ($RemoteCommand) {
                                $NetServices | ForEach-Object {
                                    $RemoteParams = @{
                                        ComputerName = $PSItem
                                        ScriptBlock  = { Write-Host "`nServer: $env:COMPUTERNAME" -ForegroundColor Cyan; Invoke-Expression -Command "$using:RemoteCommand" }
                                        ErrorAction  = 'Stop'
                                    }
                                    try {
                                        Invoke-Command @RemoteParams | Out-String
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                }
                            }
                        }
                        'KB' {
                            # #KB
                            $PortalMenuCsv | Where-Object Action -Like 'KB*' | Out-TableString
                        }
                        'KB0012657' {
                            # #KB0012657
                            # Run #7
                            Write-Host "`r`nAction: #7 - Quser: Terminal Server Logged on sessions" -ForegroundColor Cyan
                            Get-Quser -ComputerName $NetServices
                            # Run #11time
                            Write-Host "Action: #11time - Terminal Server Time" -ForegroundColor Cyan
                            Get-TSTime -ComputerName $NetServices | Out-TableString

                            # Run #11 (and if needed, select the server in question).
                            Write-Host "Action: #11nd - Terminal Server Resource Auto Query" -ForegroundColor Cyan
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            if (-not $NetServices) { $NetServices = $ComputerName }
                            Clear-Variable -Name VCAStdTSNames -ErrorAction Ignore

                            $VCAStdTSNames = Get-VcaStdTSNames -ComputerName $ComputerName | Select-Object -ExpandProperty Name
                            Get-GuestResource -ComputerName (@($NetServices) + @($VCAStdTSNames)) -Credential $script:ADCredential

                            # Run #11H (and if needed, select the server in question).
                            Write-Host "Action: #11h - Total Memory by Application by User $(Get-Date -Format g)`r`n" -ForegroundColor Cyan
                            Get-HeavyMem -ComputerName $NetServices

                            # Run #11V (and if needed, select the server in question).
                            Write-Host "Action: #11v - WOOFware Memory Use by User" -ForegroundColor Cyan
                            ############ TODO: Make sure it works with cluster sites and failed ns
                            Get-WWUserMemory -ComputerName $NetServices | Out-TableString
                            Write-Host "If users show very large sessions, recommend they sign out of WOOFware and back in to reduce memory usage.`r`n" -ForegroundColor Yellow

                            # Run #15
                            Write-Host "Action: #15 - Hospital Circuit Details" -ForegroundColor Cyan
                            Invoke-UpdateVcaCircuitsCsv
                            Get-VcaHospitalCircuits

                            # Run #20
                            Write-Host "Action: #20 - Hospital Servers in Active Directory" -ForegroundColor Cyan
                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Write-Host "`r`n > Get-ADComputer -Filter `"Name -like '$SiteAU-*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'`" -Properties OperatingSystem, IPv4Address, CanonicalName" -ForegroundColor Cyan
                            Get-VcaADComputers -AU $ComputerName | Out-TableString -Wrap
                        }
                        'lockouts' {
                            # #lockouts
                            # Show ad account lockout events by caller
                            if (-not $ADCredential) {
                                $ADCredential = Get-ADCreds
                            }
                            if ($ADCredential) {
                                Get-ADUserLockouts -Credential $ADCredential | Out-GridView -Title "#lockouts AD account lockout events by caller (last 24hrs) - Select entries to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple | Out-String
                            }
                            if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                            }
                        }
                        'mmc' {
                            # #mmc
                            #launch mmc directory
                            $MMCPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools"
                            if (Test-Path -Path $MMCPath) {
                                Start-Process "$MMCPath"
                            }
                        }
                        'ping' {
                            # #ping
                            # Check  GuestVM last reboot
                            Clear-Variable -Name SiteServers -ErrorAction Ignore
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#ping Select server to Ping Core components - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                            Invoke-Command -ComputerName $SiteServers.name  { Test-Connection google.com, phrdslp01, ladcp01, phhospdhcp1, 10.211.102.100, 10.125.105.100 }

                        }
                        'pw' {
                            # #pw
                            #generate strong password
                            Write-Host "Strong Password: " -ForegroundColor Cyan -NoNewline
                            Get-StrongPassword -Length 15 | Out-String
                        }
                        'refresh' {
                            # #refresh
                            Invoke-CachedListsRefresh
                        }
                        'robo' {
                            # #robo
                            Invoke-VcaOpsPortalUpdate
                        }
                        'san' {
                            # #san
                            if (($Cluster) -and ($ClusterSite -notlike '*vSAN*')) {
                                $ComputerNameStripped = Convert-VcaAu -AU $ComputerName -Suffix ''
                                @("$ComputerNameStripped-sanA", "$ComputerNameStripped-sanB") | ForEach-Object {
                                    Start-Process "https://$PSItem"
                                }
                            }
                        }
                        'smb' {
                            # #smb
                            #launch ns smb share
                            if (-not $NetServices) { break }
                            $NetServices | foreach-object {
                                if ($PSItem | Get-PingStatus) {
                                    Invoke-Item -Path "\\$PSItem\c$"
                                }
                            }
                        }
                        'unlock' {
                            # #unlock
                            # unlock an AD account
                            Clear-Variable -Name LockedADAccounts, LockedADAccountsSelection -ErrorAction Ignore
                            $LockedADAccounts = Search-ADAccount -LockedOut | Select-Object -Property LockedOut, Name, SamAccountName, LastLogonDate, Enabled, SID, DistinguishedName
                            $LockedADAccounts | Out-Gridview -Title "#unlock Select account to unlock - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable LockedADAccountsSelection | Out-String

                            if ($LockedADAccountsSelection) {
                                $LockedADAccountsSelection | foreach-object { Unlock-ADAccount -Identity $PSItem.SamAccountName -Credential $ADCredential -Verbose -Confirm:$true }
                            }
                        }
                        'up' {
                            # #up
                            # Check  GuestVM last reboot
                            Clear-Variable -Name SiteServers -ErrorAction Ignore
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#up Select server(s) to Check last reboot - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                            # Proceed if server was selected
                            if ($SiteServers) {
                                Get-CimInstance -ComputerName $SiteServers.name -ClassName win32_operatingsystem | Select-Object CSName, LastBootupTime | Sort-Object csname |
                                Out-GridView -Title "#up Server(s) Last Reboot - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                            }
                        }
                        'util' {
                            # #util
                            # Launch UTIL AMT Management tool
                            $UTILURL = "http://$(Convert-VcaAu -AU $ComputerName -UTIL):16992/index.htm"
                            Start-Process $UTILURL
                            Write-Output $UTILURL
                            Write-Host "`r" # add 1 blank row above the next prompt
                        }
                        '0' {
                            # #0
                            # Change host
                            Write-Host "`r`n$('#'*100)`r`n" -ForegroundColor Gray
                            Clear-Variable -Name ComputerName
                            Clear-Variable -Name Cluster -ErrorAction Ignore # Resets variables that get set in tools like #13.  The variable was cleared so that the following hospital doesn't carry over the previous hospital's values that may or may not get reset depending on if the hospital is a Cluster or not.  "-ErrorAction Ignore" was added for when users choose an incorrect hospital # right after they launch it and the variable doesn't exist yet.
                            Clear-Variable -Name ClusterSite -ErrorAction Ignore # Resets variables that get set in tools like #13.  The variable was cleared so that the following hospital doesn't carry over the previous hospital's values that may or may not get reset depending on if the hospital is a Cluster or not.  "-ErrorAction Ignore" was added for when users choose an incorrect hospital # right after they launch it and the variable doesn't exist yet.
                            Clear-Variable -Name NetServices -ErrorAction Ignore # Resets variables that get set in tools like #13.  The variable was cleared so that the following hospital doesn't carry over the previous hospital's values that may or may not get reset depending on if the hospital is a Cluster or not.  "-ErrorAction Ignore" was added for when users choose an incorrect hospital # right after they launch it and the variable doesn't exist yet.
                        }
                        '000' {
                            # #000
                            # changelog
                            Start-Process -FilePath "$PSScriptRoot\Changelog.txt"
                        }
                        '1' {
                            # #1
                            # esxcli drive - results in notepad
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential) {
                                if ((Read-Choice -Title "#1 Run ESXCLI Report on [$Computername] - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 1) -eq 1) {
                                    Get-DriveStatus -ComputerName $ComputerName -Credential $EsxiCredential -ExportPath "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Reports" | Out-String
                                    Set-EsxiCredential -Credential $EsxiCredential
                                }
                            }
                        }
                        '1l' {
                            # #1L
                            # launch reports directory
                            Invoke-Item -Path "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Reports"
                        }
                        '1p' {
                            # #1p
                            # launch reports directory
                            Invoke-Item -Path "$PSScriptRoot"
                        }
                        '1s' {
                            # #1s
                            # Search for site drive report logs
                            $DiskReports = Get-ChildItem -Path "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Reports\$ComputerName*.txt" -Recurse |
                                Select-Object -Property Name, LastWriteTime, Length, FullName | Sort-Object -Property LastWriteTime -Descending |
                                Out-GridView -Title "#1s Select drive report log(s) to launch - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            $DiskReports | Select-Object -ExpandProperty FullName | Invoke-Item
                        }
                        '2' {
                            # #2
                            # esxcli drive - results on-screen only
                            # Load Creds
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential) {
                                if ((Read-Choice -Title "#2 Run ESXCLI Report on [$Computername] - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 1) -eq 1) {
                                    Get-DriveStatus -ComputerName $ComputerName -NoLog -Credential $EsxiCredential | Out-String
                                    Set-EsxiCredential -Credential $EsxiCredential
                                }
                            }
                        }
                        '3' {
                            # #3
                            # enable ssh
                            if (Get-Module -Name VMware.VimAutomation.Core -ListAvailable) {
                                #Load Creds
                                if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                                if ($EsxiCredential) {
                                    Enable-SSH -ComputerName $ComputerName -Credential $EsxiCredential | Out-String
                                    Set-EsxiCredential -Credential $EsxiCredential
                                }
                            }
                            else {
                                Write-Warning 'VMware.VimAutomation.Core module not found.'
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name VMware.PowerCLI"
                            }
                        }
                        '4' {
                            # #4
                            # disable ssh
                            if (Get-Module -Name VMware.VimAutomation.Core -ListAvailable) {
                                #Load Creds
                                if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                                if ($EsxiCredential) {
                                    Disable-SSH -ComputerName $ComputerName -Credential $EsxiCredential | Out-String
                                    Set-EsxiCredential -Credential $EsxiCredential
                                }
                            }
                            else {
                                Write-Warning 'VMware.VimAutomation.Core module not found.'
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name VMware.PowerCLI"
                            }
                        }
                        '5' {
                            # #5
                            # launch putty
                            # Load Creds
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential) {
                                Write-Host '# ESXi SSH Commands' -ForegroundColor Cyan
                                Write-Host ''
                                Write-Host '# OS Commands:' -ForegroundColor Yellow
                                Write-Host ' hostname; date; uptime;' -ForegroundColor Gray
                                Write-Host ' esxtop;' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# Array Controller/Disk commands' -ForegroundColor Yellow
                                Write-Host ' esxcli ssacli cmd -q "ctrl slot=0 ld all show detail"' -ForegroundColor Gray
                                Write-Host ' esxcli ssacli cmd -q "ctrl slot=0 pd all show status"' -ForegroundColor Gray
                                Write-Host ' esxcli ssacli cmd -q "ctrl all show status"' -ForegroundColor Gray
                                Write-Host ' esxcli ssacli cmd -q "ctrl slot=0 ld 1 modify size=max forced"' -ForegroundColor Gray
                                Write-Host 'The command deletes the Array.  (Use with Decommissions):' -ForegroundColor Red
                                Write-Host ' esxcli ssacli cmd -q "ctrl slot=0 delete forced"' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# vSan commands' -ForegroundColor Yellow
                                Write-Host ' esxcli vsan storage list #list disk uuid' -ForegroundColor Gray
                                Write-Host ' esxcli vsan storage remove -u <UUID> #remove disk from group' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# Datastore' -ForegroundColor Yellow
                                Write-Host 'df -h' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# Filesystem Disk Info' -ForegroundColor Yellow
                                Write-Host 'vdf -h' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# Networking' -ForegroundColor Yellow
                                Write-Host 'esxcfg-nics -l' -ForegroundColor Gray
                                Write-Host 'esxcfg-vswitch -l' -ForegroundColor Gray
                                Write-Host 'esxcfg-vmknic -l' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# Misc' -ForegroundColor Yellow
                                Write-Host 'dcui' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# SNMP Get/Set' -ForegroundColor Yellow
                                Write-Host 'esxcli system snmp get' -ForegroundColor Gray
                                Write-Host 'esxcli system snmp set --communities vcapublic'-ForegroundColor Gray
                                Write-Host 'esxcli system snmp set --enable true'-ForegroundColor Gray
                                Write-Host 'esxcli system snmp get' -ForegroundColor Gray
                                Write-Host ''
                                Write-Host '# iLO: Determine iLO IP' -ForegroundColor Yellow
                                Write-Host '/opt/tools/hponcfg -w iLOconfig.txt; cat iLOconfig.txt | grep IP' -ForegroundColor Gray
                                Write-Host '# iLO: Reboot iLO' -ForegroundColor Yellow
                                Write-Host '/opt/tools/hponcfg -b' -ForegroundColor Gray
                                Write-Host ''

                                Start-Process -FilePath "$PSScriptRoot\Private\bin\putty.exe" -ArgumentList "-ssh $($EsxiCredential.UserName)@$ComputerName -pw $($EsxiCredential.GetNetworkCredential().Password)"
                                Set-EsxiCredential -Credential $EsxiCredential
                            }
                        }
                        '5s' {
                            # #5s
                            # launch sana/sanb ssh
                            if (($Cluster) -and ($ClusterSite -notlike '*vSAN*')) {
                                Write-Host '# MSA 2040' -ForegroundColor Cyan
                                Write-Host '# Midplane/Chassis Serial Number' -ForegroundColor Yellow
                                Write-Host 'show fru' -ForegroundColor Gray
                                Write-Host ''

                                $ComputerNameStripped = Convert-VcaAu -AU $ComputerName -Suffix ''
                                @("$ComputerNameStripped-sanA", "$ComputerNameStripped-sanB") | ForEach-Object {
                                    Start-Process -FilePath "$PSScriptRoot\Private\bin\putty.exe" -ArgumentList "-ssh manage@$PSItem"
                                }
                            }
                        }
                        '5i' {
                            # #5i
                            # launch ilo ssh
                            Write-Host '# iLO SSH Commands' -ForegroundColor Cyan
                            Write-Host '# Reset ilo' -ForegroundColor Yellow
                            Write-Host 'cd map1' -ForegroundColor Gray
                            Write-Host 'reset' -ForegroundColor Gray
                            Write-Host ''

                            if ((-not $Cluster) -and $NetServices) { $ServerIlo = "$ComputerName-ilo" }
                            elseif ($Cluster) {
                                $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                            }
                            if (-not $IloCredential) { $IloCredential = Get-Credential -Message 'iLO Credentials:' }

                            if ($IloCredential) {
                                $ServerIlo | ForEach-Object {
                                    Start-Process -FilePath "$PSScriptRoot\Private\bin\putty.exe" -ArgumentList "-ssh $($IloCredential.UserName)@$PSItem -pw $($IloCredential.GetNetworkCredential().Password)"
                                }
                                if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                    New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                }
                            }
                        }
                        '6' {
                            # #6
                            # launch scp
                            # Load Creds
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential -and -not $cluster) {
                                Start-Process -FilePath "$PSScriptRoot\Private\bin\winscp.exe" -ArgumentList "scp://$($EsxiCredential.UserName):$($EsxiCredential.GetNetworkCredential().Password)@$ComputerName`:/vmfs/volumes/$ComputerName/"
                                Set-EsxiCredential -Credential $EsxiCredential
                            }
                            elseif ($EsxiCredential -and $cluster) {
                                Start-Process -FilePath "$PSScriptRoot\Private\bin\winscp.exe" -ArgumentList "scp://$($EsxiCredential.UserName):$($EsxiCredential.GetNetworkCredential().Password)@$ComputerName`:/vmfs/volumes/"
                                Set-EsxiCredential -Credential $EsxiCredential
                            }
                        }
                        '7' {
                            # #7
                            # quser
                            Get-Quser -ComputerName $NetServices
                        }
                        '7g' {
                            # #7g
                            # quser state check on all hospital NSs exlcuding canada.
                            Clear-Variable -Name ServerStateSelection -ErrorAction Ignore
                            $HospitalNS = (Get-ADComputer -Filter '(Name -like "h*-ns*" -and Name -notlike "*-old" -and Name -notlike "h8*-ns*") -and OperatingSystem -like "*Server*" -and Enabled -eq $true' | Select-Object -ExpandProperty Name) -match '^h\d+-ns\d{0,2}$' | Sort-Object

                            if ($HospitalNS) {
                                $ServerState = Get-QuserStateParallel -ComputerName $HospitalNS

                                if ($ServerState) {
                                    if (-not $HospitalMaster) { $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc }
                                    if ($HospitalMaster) {
                                        $InHospitalMaster = {
                                            $InHospitalMaster_Name = $_.Name
                                            if ($_.State -eq 'The RPC server is unavailable.') {
                                                if ($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq "$(Convert-VcaAU -AU $InHospitalMaster_Name -Strip)" } )) { 'Yes' }
                                                else { '--' }
                                            }
                                        }
                                    }
                                    $PingResponse = {
                                        if ($_.State -eq 'The RPC server is unavailable.') {
                                            if ($_.Name | Get-PingStatus) { 'Yes' } else { '--' }
                                       }
                                    }
                                    #$ServerState | Out-GridView -Title "#7g Select Server(s) to open in temporary Excel document. - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerStateSelection | Out-Null
                                    $ServerStateProcessed = $ServerState | Select-Object -Property Name, State, Timestamp, SessionCount, @{n='InHospitalMaster';e=$InHospitalMaster}, @{n='PingResponse';e=$PingResponse}, Session
                                    $ServerStateProcessed | Out-GridView -Title "#7g Select Server(s) to open in temporary Excel document - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerStateSelection | Out-Null
                                    if ($ServerStateSelection) {
                                        $ServerStateSelection | Select-Object -Property Name, State, Timestamp, SessionCount, InHospitalMaster, PingResponse | Export-Excel -FreezeTopRow -BoldTopRow -AutoFilter -AutoSize -Show -Now
                                        #$ServerStateSelection | Select-Object -ExpandProperty Session | Export-Excel -FreezeTopRow -BoldTopRow -AutoFilter -AutoSize -Show -Now
                                        #$ServerStateSelection | Select-Object -Property Name | Out-TableString
                                    }
                                }
                            }
                        }
                        '7m' {
                            # #7m
                            # Terminal message
                            Clear-Variable -Name SiteServers, TerminalMessageSelection -ErrorAction Ignore

                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                Out-GridView -Title "#7m AU$(Convert-VcaAu -AU $ComputerName -Strip) - Select Server(s) to send terminal message - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null

                            if ($SiteServers) {
                                $TerminalMessage = @(
                                    "The server will be shutting down and offline for maintenance. Please save any work to prevent data loss. Thank you --IT Operations"
                                    "The server will be rebooting for maintenance. Thank you --IT Operations"
                                    "The server will be rebooting for maintenance and should return in 15-20 minutes. Thank you --IT Operations"
                                    "The server will be rebooting for maintenance. Please save any work to prevent data loss. Thank you --IT Operations"
                                    "The internet issue has been reported and is being addressed. Service should be restored shortly. Thank you --IT Operations"
                                    "The internet issue has been resolved. Thank you --IT Operations"
                                    "The server will be rebooting for maintenance. Please save any work to prevent data loss. Thank you --VCA Service Desk"
                                    "The internet issue has been reported and is being addressed. Service should be restored shortly. Thank you --VCA Service Desk"
                                    "The internet issue has been resolved. Thank you --VCA Service Desk"
                                    "--- Custom ---"
                                )
                                $TerminalMessageObj = $TerminalMessage | Select-Object -Property @{n='Message';e={$PSItem}}, @{n='Target';e={$SiteServers.Name}}
                                $TerminalMessageSelection = $TerminalMessageObj | Out-GridView -Title "#7m Select Message for Terminal Message - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

                                if ($TerminalMessageSelection.Message -eq '--- Custom ---') {
                                    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

                                    $InputBoxTitle = "Enter a custom message to send to terminal sessions"
                                    $InputBoxMessage = "Target: $($SiteServers.Name -join ', ')`r`n`nEnter a message:"
                                    $InputBoxDefault = 'The server will be rebooting for maintenance. Thank you --IT Operations'

                                    $TerminalMessageSelection.Message = [Microsoft.VisualBasic.Interaction]::InputBox($InputBoxMessage, $InputBoxTitle, $InputBoxDefault)
                                }
                                if ($TerminalMessageSelection.Message) {
                                    if ((Read-Choice -Title "#7m Are you sure you want to send a terminal message to $($SiteServers.Name -join ', ')? - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 0) -eq 1) {
                                        $SiteServers.Name | foreach-object {
                                            try {
                                                Write-Host "[$PSItem] Sending terminal message..." -ForegroundColor Cyan
                                                Write-Host " > Message: $($TerminalMessageSelection.Message)" -ForegroundColor Cyan
                                                Invoke-Command -ComputerName $PSItem { msg * $using:TerminalMessageSelection.Message } -ErrorAction Stop | Out-String
                                            }
                                            catch {
                                                Write-Warning $_.Exception.Message
                                            }
                                            Write-Host ''
                                        }
                                    } # confirm prompt
                                }
                            }
                        }
                        '7s' {
                            # #7s
                            # quser select
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#7s Select Server(s) to to run quser - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            Get-Quser -ComputerName $SiteServers.Name
                        }
                        '77' {
                            # #77
                            # session details/launch vnc
                            Clear-Variable -Name WhatUsers -ErrorAction Ignore 
                            if ((-not $Cluster) -and $NetServices) {
                                if ($NetServices | Get-PingStatus) { $WhatUsers = whatusers -ComputerName $NetServices -Credential $script:ADCredential}
                            }
                            elseif ($Cluster) {
                                Clear-Variable -Name NetservicesResults -ErrorAction Ignore
                                $NetservicesResults = $(
                                    $NetServices | ForEach-Object {
                                        if ($PSItem | Get-PingStatus) { $PSItem }
                                    }
                                )
                                $WhatUsers = whatusers -ComputerName $NetServicesResults -Credential $script:ADCredential
                            }
                            if ($WhatUsers) {
                                Clear-Variable -Name WhatUsersSelection -ErrorAction Ignore
                                $WhatUsers | Out-GridView -Title "#77 AU$(Convert-VcaAu -AU $ComputerName -Strip) - Total logged on users: $(($WhatUsers | Where-Object UserName -ne '').Count) - Select user to connect with VNC - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable WhatUsersSelection | Out-String
                                if ($WhatUsersSelection.IPAddress) {
                                    $VNCProcess = Start-Process "$PSScriptRoot\Private\bin\vncviewer.exe" -ArgumentList "$($WhatUsersSelection.IPAddress) -WarnUnencrypted=0" -PassThru
                                    Start-Sleep -Milliseconds 500

                                    # Move mouse cursor to top of screen
                                    if ($VNCProcess) {
                                        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([System.Windows.Forms.Cursor]::Position.X, 0)
                                    }
                                }
                            }
                        }
                        '77l' {
                            # #77L
                            # logoff user
                            Clear-Variable -Name WhatUsers -ErrorAction Ignore
                            if ((-not $Cluster) -and $NetServices) {
                                if ($NetServices | Get-PingStatus) { $WhatUsers = whatusers -ComputerName $NetServices -Credential $script:ADCredential}
                            }
                            elseif ($Cluster) {
                                Clear-Variable -Name NetservicesResults -ErrorAction Ignore
                                $NetservicesResults = $(
                                    $NetServices | ForEach-Object {
                                        if ($PSItem | Get-PingStatus) { $PSItem }
                                    }
                                )
                                $WhatUsers = whatusers -ComputerName $NetServicesResults -Credential $script:ADCredential
                            }

                            if ($WhatUsers) {
                                Clear-Variable -Name WhatUsersSelection, QuserCount -ErrorAction Ignore
                                $WhatUsers | Out-GridView -Title "#77L AU$(Convert-VcaAu -AU $ComputerName -Strip) - Total logged on users: $(($WhatUsers | Where-Object UserName -ne '').Count) - Select user(s) to log off - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable WhatUsersSelection |
                                    Where-Object { $_.UserName -and $_.SessionId -ne '99999' } | Out-ListString
                                $WhatUsersSelection = $WhatUsersSelection | Where-Object { $_.UserName -and $_.SessionId -ne '99999' }
                                if ($WhatUsersSelection.UserName) {
                                    $WhatUsersSelection | ForEach-Object {
                                        $WhatUsersSelection_Item = $PSItem
                                        if ($PSItem.UserName) {
                                            #Start-Process -FilePath logoff.exe -ArgumentList "$($PSItem.SessionId) /SERVER:$($PSItem.Computer)" -NoNewWindow -Wait | Out-String
                                            Write-Host "[$($PSItem.Computer)] Logging off: $($PSItem.UserName) (ID:$($PSItem.SessionId))" -ForegroundColor Cyan
                                            try {
                                                Stop-TSSession -ComputerName $PSItem.Computer -Id $PSItem.SessionId -Force -ErrorAction Stop
                                            }
                                            catch [System.Management.Automation.MethodInvocationException] {
                                                Write-Warning "[$($WhatUsersSelection_Item.Computer)] Session ID:$($WhatUsersSelection_Item.SessionId) is invalid or no longer exists"
                                            }
                                            catch {
                                                Write-Warning "[$($WhatUsersSelection_Item.Computer)] $($_.Exception.Message)"
                                            }
                                        }
                                    }
                                    $WhatUsersSelection.Computer | Select-Object -Unique | ForEach-Object {
                                        Write-Host "`r`nServer: $PSItem ($([System.Net.Dns]::GetHostAddresses($PSItem).IPAddressToString))`n"
                                        quser.exe "/SERVER:$PSItem" | Tee-Object -Variable QuserCount
                                        Write-Host "`r`nCount of users: $(@($QuserCount).count - 1)`r`n"
                                    }
                                    Write-Host "Users may still be in the process of being logged out, please query site again to verify.`r`n" -ForegroundColor Cyan
                                }
                            }
                        }
                        '77s' {
                            # #77s
                            # session details/launch vnc
                            Clear-Variable -Name WhatUsers -ErrorAction Ignore
                            if ((-not $Cluster) -and $NetServices) {
                                if ($NetServices | Get-PingStatus) { $WhatUsers = whatusers -ComputerName $NetServices -Credential $script:ADCredential}
                            }
                            elseif ($Cluster) {
                                Clear-Variable -Name NetservicesResults -ErrorAction Ignore
                                $NetservicesResults = $(
                                    $NetServices | ForEach-Object {
                                        if ($PSItem | Get-PingStatus) { $PSItem }
                                    }
                                )
                                $WhatUsers = whatusers -ComputerName $NetServicesResults -Credential $script:ADCredential
                            }
                            if ($WhatUsers) {
                                Clear-Variable -Name WhatUsersSelection -ErrorAction Ignore
                                $WhatUsers | Out-GridView -Title "#77s AU$(Convert-VcaAu -AU $ComputerName -Strip) - Total logged on users: $(($WhatUsers | Where-Object UserName -ne '').Count) - Select user to shadow. - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable WhatUsersSelection | Out-String
                                if ($WhatUsersSelection.IPAddress) {
                                    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$($WhatUsersSelection.Computer) /shadow:$($WhatUsersSelection.SessionId) /control"
                                }
                            }
                        }
                        '8' {
                            # #8
                            # headcount per Active Directory
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteUsers -ErrorAction Ignore
                                $SiteUsers = Get-VCAHeadcount -AU $ComputerName
                                Write-Host ''
                                if ($SiteUsers.UserCount) {
                                    Write-Host "$SiteAUNumber - Active Directory User Count: $($SiteUsers.UserCount)`r`n" -ForegroundColor Cyan
                                }

                                $ADUserParams = @{
                                    Property = @(
                                        'Name'
                                        'Title'
                                        'OfficePhone'
                                        'Office'
                                        'Department'
                                        'EmailAddress'
                                        'StreetAddress'
                                        'City'
                                        'State'
                                        'PostalCode'
                                        'SID'
                                        'Created'
                                        @{n='extensionAttribute3';e={if ($_.extensionAttribute3) {$_.extensionAttribute3 -replace "\d","*"}}}
                                    )
                                }
                                $SiteUsers.Users | Select-Object @ADUserParams | Sort-Object -Property Name, Title |
                                Out-GridView -Title "#8 $SiteAUNumber - Total user accounts: $($SiteUsers.UserCount) - Select user(s) to diplay in console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteUserSelection | Out-Null

                                if ($SiteUserSelection.extensionAttribute3) {
                                    Write-Host "*** Cell phone numbers are private information and should ONLY be used for emergencies" -ForegroundColor Magenta
                                    Write-Host "    DO NOT distribute or disclose without the owners permission ***`r`n" -ForegroundColor Magenta
                                }
                                $ADProperties = @(
                                    'Name'
                                    'Title'
                                    'OfficePhone'
                                    'Office'
                                    'Department'
                                    'EmailAddress'
                                    'StreetAddress'
                                    'City'
                                    'State'
                                    'PostalCode'
                                    'SID'
                                    'Created'
                                    'extensionAttribute3'
                                )
                                $SiteUsers.Users | Where-object {$_.SID -in $SiteUserSelection.SID} | Select-Object -Property $ADProperties | Out-ListString -NoNewLine
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '8t' {
                            # #8t
                            # Open ServiceNow ticket via filtered Actiev Directory list of hospital users
                            Clear-Variable -Name UserTicket -ErrorAction Ignore
                            $SiteUsers = Get-VCAHeadcount -AU $ComputerName
                            $ADUserParams = @{
                                Property = @(
                                    'Name'
                                    'Title'
                                    'OfficePhone'
                                    'Office'
                                    'Department'
                                    'EmailAddress'
                                    'StreetAddress'
                                    'City'
                                    'State'
                                    'PostalCode'
                                    'SID'
                                    'Created'
                                )
                            }
                            $SiteUsers.Users | Select-Object @ADUserParams | Sort-Object -Property Name, Title |
                            Out-GridView -Title "#8t AU$($SiteUsers.AU) - Total user accounts: $($SiteUsers.UserCount) - Select user to generate ticket - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable UserTicket | Out-String

                            if ($UserTicket) {
                                Invoke-SnowGui -ComputerName $ComputerName -ImpactedUser $UserTicket.EmailAddress -ImpactedUserList $SiteUsers.Users.EmailAddress -HospitalMaster $HospitalMaster
                            }
                        }
                        '9' {
                            # #9
                            # launch ilo page
                            #if ($ComputerName -match '-vm$|-vm\d{1,2}') {
                            if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$') {
                                Start-Process "https://$ComputerName-ilo".Replace('.vcaantech.com','')
                            }
                            else {
                                Start-Process "https://$ComputerName"
                            }
                        }
                        '10' {
                            # #10
                            # launch esxi webui
                            Start-Process "https://$ComputerName/ui"
                        }
                        '11?' {
                            # #11?
                            $PortalMenuCsv | Where-Object Action -Like '11*' | Out-TableString
                        }
                        '11' {
                            # #11
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }

                            $SiteServers = Select-VcaSite -AU $ComputerName -Title "#11 Select Server(s) to query resources - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            if ($SiteServers.Name) {
                                Get-GuestResource -ComputerName $SiteServers.Name -Credential $script:ADCredential
                            }
                            elseif ($SingleHost) {
                                Get-GuestResource -ComputerName $ComputerName -Credential $script:ADCredential
                            }
                        }
                        '11d' {
                            # #11d
                            # db memory, drive, cpu usage and uptime
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            Get-GuestResource -ComputerName $(Convert-VcaAU -AU $ComputerName -Suffix '-db') -Credential $script:ADCredential
                        }
                        '11dl' {
                            # #11dl
                            # check for disk latency
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore

                                $SiteServers = Select-VcaSite -AU $ComputerName -Title "#11dl Select Server(s) to query resources - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

                                if ($SiteServers) {
                                    if (($SiteServers -match '-dc') -and (-not $ADCredential)) { $ADCredential = Get-ADCreds }
                                    Get-DiskUsage -ComputerName $SiteServers.Name -Credential $ADCredential | Out-GridView -Title "#11dl Select Drive(s) to check for latency - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable SiteDrives | Out-Null
                                    
                                    if ($siteDrives.ComputerName -eq $null ) {
                                        Write-Host "`r`nDisk selection was cancelled.  Aborting.`r`n" -ForegroundColor Cyan
                                        } else {
                                        Invoke-Command -ComputerName $SiteDrives.ComputerName -scriptblock {
                                        $cntr = 0 

                                        $Counters = @(("\\$($using:SiteDrives.ComputerName)\LogicalDisk($($using:SiteDrives.Name))\Avg. disk sec/transfer"))
                                        
                                        $disksectransfer = Get-Counter -Counter $Counters -MaxSamples 1 
                                        $avg = $($disksectransfer.CounterSamples | Select-Object CookedValue).CookedValue

                                        Get-Counter -Counter $Counters -SampleInterval 2 -MaxSamples 30 | ForEach-Object {
                                        $_.CounterSamples | ForEach-Object {
                                        [pscustomobject]@{
                                            TimeStamp = $_.TimeStamp
                                            Path = $_.Path
                                            Value = ([Math]::Round($_.CookedValue, 5))
                                                turn = $cntr = $cntr +1
                                                running_avg = [Math]::Round(($avg = (($_.CookedValue + $avg) / 2)), 5)  
                                            
                                                }   
                                            }
                                        }
                                    } | Format-Table

                                    write-host "Final_Running_Average: $([Math]::Round( $avg, 5)) sec/transfer`n"
                                        
                                    if ($avg -gt 0.01) {
                                        Write-Host "There ARE indications of slow I/O performance on your system - avg greater than 0.01"
                                        } else {
                                        Write-Host "There is NO indication of slow I/O performance on your system -  avg less than 0.01"
                                        }
                                    }   
                                }
                            }
                        }
                        '11f' {
                            # #11f
                            # flush memory (runtime broker)
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Invoke-Command -ComputerName $NetServices -ScriptBlock { Get-Process -Name RuntimeBroker | Stop-Process -Verbose } | Out-String
                            }
                            Write-Host ''
                        }
                        '11h' {
                            # #11h
                            # user session heavy hitter memory Teams,Chrome,WW,RuntimeBroker
                            Write-Host "`r`nTotal Memory by Application by User $(Get-Date -Format g)`r`n" -ForegroundColor Cyan
                            Get-HeavyMem -ComputerName $NetServices
                        }
                        '11ht' {
                            # #11ht
                            # user session heavy hitter memory Teams,Chrome,WW,RuntimeBroker
                            Write-Host "`r`nTotal Memory by Application $(Get-Date -Format g)" -ForegroundColor Cyan
                            Get-HeavyMemTot -NetServices $NetServices
                        }
                        '11m' {
                            # #11m
                            # user session memory
                            Write-Host "`r`nTotal Memory by User  $(Get-Date -Format g)" -ForegroundColor Cyan
                            Get-UserMemory -NetServices $NetServices | Out-String
                        }
                        '11n' {
                            # #11n
                            # ns memory, drive, cpu usage and uptime
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            if (-not $NetServices) { $NetServices = $ComputerName }
                            Get-GuestResource -ComputerName $NetServices -Credential $script:ADCredential
                        }
                        '11nd' {
                            # #11nd
                            # ns/db memory, drive, cpu usage and uptime
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            if (-not $NetServices) { $NetServices = $ComputerName }
                            Clear-Variable -Name VCAStdTSNames -ErrorAction Ignore

                            $VCAStdTSNames = Get-VcaStdTSNames -ComputerName $ComputerName | Select-Object -ExpandProperty Name
                            Get-GuestResource -ComputerName (@($NetServices) + @($VCAStdTSNames)) -Credential $script:ADCredential
                        }
                        '11net' {
                            # #11net
                            $SiteServers = Select-VcaSite -AU $ComputerName -Title "#11net Select Server(s) to query network adapters - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            if ($SiteServers.Name) {
                                Get-WindowsNetwork -ComputerName $SiteServers.Name -Credential $ADCredential | Out-GridView -Title "#11net $SiteAuNumber - Windows Network Adapters - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple | Out-ListString
                            }
                        }
                        '11p' {
                            # #11p
                            # View process memory use
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Clear-Variable -Name SiteServers, ProcessSelection -ErrorAction Ignore

                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                    Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                        Out-GridView -Title #11p Select Server(s) to check - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null

                                if ($SiteServers) {
                                    $Processes = Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock {
                                        Get-Process -IncludeUserName | Select-Object -Property PSComputerName, Username, ProcessName, WorkingSet, Id, Handle, Path
                                    } -ErrorAction SilentlyContinue
                                    $Processes | Select-Object -Property PSComputerName, Username, ProcessName, WorkingSet, Id, Handle, Path |
                                        Sort-Object -Property PSComputerName, ProcessName |
                                        Out-GridView -Title "#11p Select process for memory use summary - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable ProcessSelection | Out-Null
                                    $ProcessMemoryResults = $Processes | Where-Object { $PSItem.ProcessName -eq $ProcessSelection.ProcessName } |
                                        Group-Object -Property Username | ForEach-Object {
                                        [PSCustomObject]@{
                                            MemoryMB     = '{0:N2}' -f (($PSItem.Group.WorkingSet | Measure-Object -Sum).Sum / 1MB)
                                            Username     = $PSItem.Name
                                            Computer     = @($PSItem.Group.PSComputerName)[0]
                                            ProcessName  = @($PSItem.Group.ProcessName)[0]
                                            ProcessCount = $PSItem.Count
                                        }
                                    }
                                    $ProcessMemoryResults | Sort-Object -Property Computer | Group-Object -Property Computer | ForEach-Object {
                                        $ForeachPipe = $PSItem
                                        ($PSItem.Group | Select-Object -Property Username, MemoryMB, @{n='Computer';e={$ForeachPipe.Name}}, ProcessName, ProcessCount |
                                            Sort-Object -Property MemoryMB -Descending | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''
                                        Write-Host "[$($PSItem.Name)] Total Process memory use: $(($PSItem.Group.MemoryMB | Measure-Object -Sum).Sum) MB"
                                    }
                                }
                            }
                            Write-Host ''
                        }
                        '11pk' {
                            # #11pk
                            # Kill process in use
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Clear-Variable -Name SiteServers, ProcessSelection -ErrorAction Ignore
                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                    Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                    Out-GridView -Title "#11pk Select Server(s) to check - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null
                                if ($SiteServers) {
                                    $Processes = Invoke-Command -ComputerName $SiteServers.Name -Credential $script:ADCredential -ScriptBlock {
                                        Get-Process -IncludeUserName | Select-Object -Property PSComputerName, Username, ProcessName, Status, WorkingSet, Id, Handle, Path
                                        $Processes | Select-Object -Property PSComputerName, Username, ProcessName, ID, Status , WorkingSet, Handle, Path |
                                            Sort-Object -Property PSComputerName, ProcessName
                                    } | Out-GridView -Title "#11pk Select Process(s) to TASKKILL *** NOTE:This cannot be undone *** - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ProcessSelection | Out-Null

                                    Invoke-Command -ComputerName $SiteServers.Name -Credential $script:ADCredential -ScriptBlock {
                                        $Using:ProcessSelection | ForEach-Object {
                                            taskkill /f /PID $PSitem.id
                                        }
                                    }
                                }
                            }
                            Write-Host ''
                        }
                        '11pkwmi' {
                            # #11pkiwmi
                            # Kill process in use
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Clear-Variable -Name SiteServers, ProcessSelection -ErrorAction Ignore
                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                    Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                    Out-GridView -Title "#11pkwmi Select Server(s) to check - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null
                                if ($SiteServers) {
                                    $Processes = Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock {
                                        Get-Process -IncludeUserName -Name WmiPrvSE| Select-Object -Property PSComputerName, Username, ProcessName, Status, WorkingSet, Id, Handle, Path
                                        $Processes | Select-Object -Property PSComputerName, Username, ProcessName, ID, Status , WorkingSet, Handle, Path |
                                        Sort-Object -Property PSComputerName, ProcessName } |
                                        Out-GridView -Title "#11pkwmi Select Process(s) to TASKKILL *** NOTE:This cannot be undone *** - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ProcessSelection | Out-Null

                                    Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock {
                                        $Using:ProcessSelection | ForEach-Object {
                                            taskkill /f /PID $PSitem.id
                                        }
                                    }
                                }
                            }
                            Write-Host ''
                        }
                        '11r' {
                            # #11r
                            #if (-not (Get-Module -Name RemoteDesktop)) { Import-Module -Name RemoteDesktop -WarningAction Ignore 3>$null }
                            if (-not $Cluster -and @($NetServices).count -eq 1) {
                                try {
                                    Get-RDSConnectionConfig -ComputerName $NetServices -ErrorAction Stop | Out-ListString
                                    Write-Host "* Elevation is not needed, please ignore any warnings`r`n" -ForegroundColor Cyan
                                }
                                catch {
                                    Write-Warning $_.Exception.Message
                                }
                                break
                            }
                            if ($Cluster -or @($NetServices).count -ge 2) {
                                Clear-Variable -Name FileServer -ErrorAction Ignore
                                $FileServer = (Get-ADComputer -Filter "Name -like '$(Convert-VcaAU -AU $ComputerName -Suffix -fs)*'").Name
                                $FileServer | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        try {
                                            Get-RDSConnectionConfig -ComputerName $PSItem -ErrorAction Stop | Out-ListString
                                        }
                                        catch {
                                            Write-Warning $_.Exception.Message
                                        }
                                    }
                                    else {
                                        Write-Warning "$PSItem Connection failed"
                                    }
                                }
                            }
                            Write-Host "* Elevation is not needed, please ignore any warnings`r`n" -ForegroundColor Cyan
                        }
                        '11s' {
                            # #11s
                            if (-not (Get-Module -Name RemoteDesktop)) { Import-Module -Name RemoteDesktop -WarningAction Ignore 3>$null }

                            Clear-Variable -Name Results, MenuSelection, ConnectionsEnabledIndex, RDSHostsTo*, ConnectionBroker, CollectionName -ErrorAction Ignore
                            if (-not $Cluster -and @($NetServices).count -eq 1) {
                                try {
                                    $ConnectionBroker = [System.Net.Dns]::GetHostEntry($NetServices).HostName
                                    Write-Host "`r`n > `$CollectionName = (Get-RDSessionCollection -ConnectionBroker $ConnectionBroker).CollectionName" -ForegroundColor Cyan
                                    $CollectionName = (Get-RDSessionCollection -ConnectionBroker $ConnectionBroker).CollectionName
                                    Write-Host " > Get-RDSessionHost -CollectionName `$CollectionName -ConnectionBroker $ConnectionBroker" -ForegroundColor Cyan

                                    $Results = Get-RDSessionHost -CollectionName $CollectionName -ConnectionBroker $ConnectionBroker
                                    Write-Host "`r`n* Elevation is not needed, please ignore any warnings" -ForegroundColor Cyan
                                    $Results | Out-TableString
                                }
                                catch {
                                    Write-Warning $_.Exception.Message
                                }
                                break
                            }
                            if ($Cluster -or @($NetServices).count -ge 2) {
                                ($SiteADComputers | Where-Object Name -like "*-fs*").Name | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        try {
                                            $ConnectionBroker = [System.Net.Dns]::GetHostEntry($PSItem).HostName
                                            Write-Host "`r`n > `$CollectionName = (Get-RDSessionCollection -ConnectionBroker $ConnectionBroker).CollectionName" -ForegroundColor Cyan
                                            $CollectionName = (Get-RDSessionCollection -ConnectionBroker $ConnectionBroker).CollectionName
                                            Write-Host " > Get-RDSessionHost -CollectionName `$CollectionName -ConnectionBroker $ConnectionBroker" -ForegroundColor Cyan

                                            $Results = Get-RDSessionHost -CollectionName $CollectionName -ConnectionBroker $ConnectionBroker -ErrorAction Stop |
                                                Sort-Object -Property SessionHost
                                            Write-Host "`r`n* Elevation is not needed, please ignore any warnings" -ForegroundColor Cyan
                                            $Results | Out-TableString

                                            # Get item index of RDS allowed servers for show-menu initial selection
                                            $ConnectionsEnabledIndex = (0..(@($Results).Count-1)) | Where-Object { $Results.NewConnectionAllowed[$_] -eq 'Yes' }

                                            Write-Host "Use spacebar to select or deselect RDS server(s) for allowed connections:`r`n" -ForegroundColor Cyan
                                            # Show menu selection
                                            $MenuSelection = Show-Menu -MenuItems $Results -MultiSelect -MenuItemFormatter { ($Args).SessionHost } -InitialSelection $ConnectionsEnabledIndex -ItemFocusColor Cyan
                                            $MenuSelected = $MenuSelection | Where-Object Selected -eq $true | Select-Object -ExpandProperty MenuItem
                                            $MenuUnSelected = $MenuSelection | Where-Object Selected -eq $false | Select-Object -ExpandProperty MenuItem

                                            # Preview selection
                                            #$MenuSelection | Out-TableString

                                            [System.Collections.ArrayList]$RDSPendingChanges = @()
                                            $Results | ForEach-Object {
                                                # Mark pending task to allow new connections if disabled
                                                if ($_.SessionHost -in $MenuSelected.SessionHost) {
                                                    if ($_.NewConnectionAllowed -eq 'No' -or $_.NewConnectionAllowed -eq 'NotUntilReboot') {
                                                        $RDSPendingChanges.Add(($_ | Select-Object -Property CollectionName, SessionHost, NewConnectionAllowed, @{n='PendingTask';e={'Enable'}})) | Out-Null
                                                    }
                                                }
                                                # Mark pending task to disable new connections if not already disabled
                                                elseif ($_.SessionHost -in $MenuUnSelected.SessionHost) {
                                                    if ($_.NewConnectionAllowed -eq 'Yes') {
                                                        $RDSPendingChanges.Add(($_ | Select-Object -Property CollectionName, SessionHost, NewConnectionAllowed, @{n='PendingTask';e={'Disable'}})) | Out-Null
                                                    }
                                                }
                                            }
                                            if ($RDSPendingChanges) {
                                                Write-Host "`r`nRDS New Connection Allowed Pending Changes:" -ForegroundColor Cyan
                                                $RDSPendingChanges | Out-TableString

                                                $RDSHostsToEnable = ($RDSPendingChanges | Where-Object PendingTask -EQ 'Enable').SessionHost
                                                $RDSHostsToDisable = ($RDSPendingChanges | Where-Object PendingTask -EQ 'Disable').SessionHost
                                                if ($RDSHostsToEnable -or $RDSHostsToDisable) {
                                                    if ((Read-Choice -Title "#11s Are you sure you want to make selected changes? - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 0) -eq 1) {
                                                        if ($RDSHostsToEnable) {
                                                            Write-Host "`r`nEnabling New RDS Connections: $($RDSHostsToEnable -join ', ')" -ForegroundColor Cyan
                                                            Set-RDSessionHost -SessionHost $RDSHostsToEnable -NewConnectionAllowed Yes -ConnectionBroker $ConnectionBroker -Verbose
                                                        }
                                                        if ($RDSHostsToDisable) {
                                                            Write-Host "`r`nDisabling New RDS Connections: $($RDSHostsToDisable -join ', ')" -ForegroundColor Cyan
                                                            Set-RDSessionHost -SessionHost $RDSHostsToDisable -NewConnectionAllowed No -ConnectionBroker $ConnectionBroker -Verbose
                                                        }
                                                        Get-RDSessionHost -CollectionName $CollectionName -ConnectionBroker $ConnectionBroker | Out-TableString
                                                    }
                                                }
                                            }
                                            else {
                                                Write-Host "`r`nNo changes to make.`r`n" -ForegroundColor Cyan
                                            }
                                        }
                                        catch {
                                            Write-Warning $_.Exception.Message
                                        }
                                    }
                                    else {
                                        Write-Warning "$PSItem Connection failed"
                                    }
                                }
                            }
                        }
                        '11sm' {
                            # #11sm
                            # ns memory, drive, cpu usage and uptime
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            Get-GuestResource -ComputerName $(Convert-VcaAU -AU $ComputerName -Suffix '-smpacs') -Credential $script:ADCredential
                        }
                        '11time' {
                            # #11time
                            # Retreive date/time with timezone from ns
                            Get-TSTime -ComputerName $NetServices | Out-TableString
                        }
                        '11u' {
                            # #11u
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }
                            $SiteServers = Select-VcaSite -AU $ComputerName -Title "#11u Select Server(s) to query resources - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            if ($SiteServers.Name) {
                                Write-Host "`r`n CPU Utilization and Currently Available Memory -- 30 Second Sample`r`n" -ForegroundColor Cyan
                                $SiteServers.Name | Start-RSJob -Name WinCpuJobs -VariablesToImport ADCredential -Throttle 12 -ScriptBlock {
                                        Invoke-Command -ComputerName $_ -Credential $ADCredential -ScriptBlock {
                                            $totalRam = (Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).Sum
                                            For ($i = 1; $i -le 10; $i++) {
                                                Clear-Variable -Name Results -ErrorAction Ignore
                                                $date = Get-Date -Format "h:mm:ss tt"
                                                $cpuTime = (Get-Counter -Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
                                                $availMem = (Get-Counter -Counter '\Memory\Available MBytes').CounterSamples.CookedValue

                                                [PSCustomObject]@{
                                                    Name         = $env:computername
                                                    Timestamp    = $date
                                                    'CpuLoad(%)' = [decimal]$cpuTime.ToString("N2")
                                                    MemFreeMB    = $availMem.ToString("N0") + ' (' + (104857600 * $availMem / $totalRam).ToString("#,0.0") + '%)'
                                                }
                                                Start-Sleep -Milliseconds 1500
                                            }
                                        }
                                } | Out-Null
                                $WinCpuResults = Get-RSJob -Name WinCpuJobs | Wait-RSJob -ShowProgress -Timeout 300 | Receive-RSJob
                                Get-RSJob -Name WinCpuJobs | Remove-RSJob -Force

                                $WinCpuResults | Group-Object -Property Name | ForEach-Object {
                                    $_.Group  | Select-Object -Property Name, Timestamp, 'CpuLoad(%)', MemFreeMB | Out-TableString -NoNewLine
                                }
                            }
                        }
                       '11v' {
                            # #11v
                            # View WOOFware memory use
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Clear-Variable -Name WWMemoryResults -ErrorAction Ignore

                                $WWMemoryResults = Get-WWUserMemory -ComputerName $NetServices
                                $WWMemoryResults | Out-TableString
                                Write-Host 'If users show very large sessions, recommend they sign out of WOOFware and back in to reduce memory usage.' -ForegroundColor Yellow

                                if ($(Read-Choice -Title "#11v WOOFware Memory Sessions - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message 'Copy to clipboard?') -eq 1) {
                                    $WWMemoryResults | Select-Object -Property MemoryMB, Username, Computer, ProcessName, StartTime, ID |
                                        Convertto-HTML | Set-Clipboard -AsHtml
                                }
                            }
                            Write-Host ''
                        }
                        '11w' {
                            # #11w
                            # download and run wiztree on target machine
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#11w Select Windows machine to query disk space usage - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

                                if ($SiteServers.Name) {
                                    Write-Host "`r`n[$($SiteServers.Name)] Connecting to machine" -ForegroundColor Cyan
                                    $PSSession = New-PSSession -ComputerName $SiteServers.Name -Credential $ADCredential

                                    # Copy executable to remote machine if it doesn't exist
                                    # MD5 Hash: 50A40274FFE963E1F214F9F19746E29E
                                    $WizTreePath = 'C:\temp\WizTree64.exe'
                                    if (-not (Invoke-Command -Session $PSSession { Test-Path -Path $using:WizTreePath })) {
                                        Invoke-Command -Session $PSSession {
                                            if (-not (Test-Path -Path 'C:\temp\')) { New-Item -ItemType Directory -Path 'C:\temp\' | Out-Null }
                                        }
                                        try {
                                            Write-Host "[$($PSSession.ComputerName)] Copying over WizTree64.exe" -ForegroundColor Cyan
                                            Copy-Item -Path "$PSScriptRoot\Private\bin\WizTree64.exe" -Destination 'C:\temp\WizTree64.exe' -ToSession $PSSession -ErrorAction Stop
                                        }
                                        catch {
                                            Write-Warning $_.Exception.Message
                                        }
                                    }
                                    if ((Invoke-Command -Session $PSSession { Test-Path -Path $using:WizTreePath })) {
                                        $SiteServers | foreach-object {
                                            try {
                                                # wiztree cmdline https://diskanalyzer.com/guide
                                                $sb = {
                                                    Write-Host "[$env:COMPUTERNAME] Running WizTree64.exe`r`n" -ForegroundColor Cyan
                                                    Start-Process -FilePath $using:WizTreePath -ArgumentList 'C: /export="C:\temp\wiztree.csv" /treemapimagefile="C:\temp\wiztree.png" /admin=1 /exportfolders=1 /exportfiles=0 /sortby=1 /exportmaxdepth=5' -Wait
                                                    Get-Content -Path 'C:\temp\wiztree.csv' -TotalCount 10000 -ErrorAction Stop | Select-Object -Skip 1 | ConvertFrom-Csv |
                                                        Select-Object -Property 'File Name', @{n='SizeMB';e={$_.Size / 1MB}}, Modified, Attributes, PSComputerName
                                                }
                                                $WizTreeResults = Invoke-Command -Session $PSSession -ScriptBlock $sb -ErrorAction Stop |
                                                    Select-Object -Property 'File Name', SizeMB, Modified, Attributes, PSComputerName

                                                Copy-Item -Path "C:\temp\wiztree.png" -Destination 'C:\temp' -FromSession $PSSession
                                                if (Test-Path -Path 'C:\temp\wiztree.png') { Invoke-Item -Path 'C:\temp\wiztree.png' }

                                                $WizTreeResults | Out-GridView -Title "#11w [$($SiteServers.Name)] WizTree 10,000 largest directories on C:\ - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                                            }
                                            catch {
                                                Write-Warning $_.Exception.Message
                                            }
                                        }
                                    }
                                    if ($PSSession) { Remove-PSSession -Session $PSSession }
                                }
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '12' {
                            # #12
                            # expand windows drive
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore

                                Get-ADComputer -Filter "Name -like '$(Convert-VcaAu -AU $ComputerName -Suffix '')-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties IPv4Address, OperatingSystem |
                                    Where-Object IPv4Address -ne $null | Select-Object -Property Name, IPv4Address, OperatingSystem | Sort-Object -Property Name |
                                    Out-GridView -Title "#12 Select Server(s) to expand - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null

                                if ($SiteServers) {
                                    if (($SiteServers -match '-dc') -and (-not $ADCredential)) { $ADCredential = Get-ADCreds }
                                    Get-DiskUsage -ComputerName $SiteServers.Name -Credential $ADCredential | Out-GridView -Title "#12 Select Drive(s) to expand - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteDrives | Out-Null
                                    $SiteDrives | Out-TableString

                                    $SiteDrives | foreach-object {
                                        $ExpandDriveParams = @{ ComputerName = $PSItem.ComputerName }
                                        if ($ADCredential) { $ExpandDriveParams.Add('Credential', $ADCredential) }
                                        Invoke-Command @ExpandDriveParams -ScriptBlock { 'rescan', "sel vol $($using:PSItem.Name -replace ':','')", 'extend' | diskpart }

                                        if ($SiteServers.count -ge 2) { Write-Host "$('-'*55)`r" -ForegroundColor Cyan }
                                    }
                                    Get-DiskUsage -ComputerName $SiteDrives.ComputerName -Credential $ADCredential | Out-TableString
                                }
                            }
                            else {
                                if ($NetServices) {
                                    Get-DiskUsage -ComputerName $NetServices -Credential $ADCredential | Out-TableString
                                    Invoke-Command -ScriptBlock { 'rescan', 'sel vol c', 'extend' | diskpart } -ComputerName $NetServices -Credential $ADCredential
                                    Get-DiskUsage -ComputerName $NetServices -Credential $ADCredential | Out-TableString
                                }
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                            if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                            }
                            Write-Host ''
                        }
                        '13' {
                            # #13
                            # pinginfoview
                            Clear-Variable -Name SiteHostnames -ErrorAction Ignore
                            Write-Host "PingInfoView Instructions:" -ForegroundColor Cyan
                            Write-Host "Find the 'Order' column and make it the 1st column.  Next sort 'Order' from low to high." -ForegroundColor Yellow
                            Write-Host "...Doing so lists devices in the correct boot/connectivity order." -ForegroundColor Yellow
                            Write-Host "Highlight the device you are interested in.  Take a screenshot of the entire application for tickets." -ForegroundColor Yellow
                            Write-Host ''
                            Write-Host "Note: Cloud-based RDP Cluster sites may have hidden -vm* & -vm*-ilo hosts.  Please add them manually as needed to PingInfoView." -ForegroundColor Red
                            Write-Host ''

                            if ($ComputerName -match "^h\d{1,6}-vm") {
                                $SiteHostnames = Get-VcaSiteHostname -ComputerName $ComputerName -Cluster $Cluster -ClusterSite $ClusterSite -NetServices $NetServices
                            }
                            else { $SiteHostnames = $ComputerName }

                            Start-PingInfoView -ComputerName $SiteHostnames
                        }
                        '13p' {
                            # #13p
                            # ping plotter
                            Clear-Variable -Name PingPlotterSite -ErrorAction Ignore

                            $PingPlotterSite = Convert-VcaAu -AU $ComputerName -Suffix '-gw'
                            if ($PingPlotterSite) {
                                $PingPlotterPID = (Start-Process "$PSScriptRoot\Private\bin\PingPlotter 5\PingPlotter.exe" -PassThru).Id
                                Start-Sleep -Milliseconds 3500
                                $PingPlotterWindow = Select-UIElement -PID $PingPlotterPID
                                $PingPlotterWindow | Select-UIElement -ControlType 'Edit' -AutomationId 'TextBox' -Recurse |
                                    Set-UIText -Text $PingPlotterSite -Passthru | Send-UIKeys -Keys '{enter}' | Out-Null
                            }
                        }
                        '13pc' {
                            # #13pc
                            # gping
                                Start-Process -FilePath "$PSScriptRoot\Private\bin\gping.exe" -ArgumentList "$(Convert-VcaAu -AU $ComputerName -Suffix '-gw') $(Convert-VcaAu -AU $ComputerName -Suffix '-ns') -b 60"
                        }
                        '14' {
                            # #14
                            # Hospital Master
                            if (-not $HospitalMaster) {
                                $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
                            }
                            if ($HospitalMaster) {
                                $HospitalInfo = $HospitalMaster.Where( {
                                        $PSItem.'Hospital Number' -eq "$(Convert-VcaAU -AU $ComputerName -Strip)" } )

                                if ($HospitalInfo) {
                                    Write-Host "`r`nLocation:"
                                    Write-Host "$($HospitalInfo.'Operating Name') #$($HospitalInfo.'Hospital Number')"
                                    Write-Host "$($HospitalInfo.Address)"
                                    Write-Host "$($HospitalInfo.City), $($HospitalInfo.St) $($HospitalInfo.Zip)"
                                    Write-Host ''
                                    Write-Host 'VCA Site Contact:'
                                    Write-Host "$($HospitalInfo.'Hospital Manager'), $($HospitalInfo.'Hospital Manager Email')"
                                    Write-Host "$($HospitalInfo.Phone)"
                                    Write-Host ''
                                    Write-Host 'Misc. info:'
                                    Write-Host 'Time Zone              :'"$($HospitalInfo.'Time zone')"
                                    Write-Host 'URL                    :'"$($HospitalInfo.GPURL)"
                                    Write-Host 'Back Line              :'"$($HospitalInfo.'Back Line')"
                                    Write-Host 'System Conversion Date :'"$($HospitalInfo.'System Conversion Date')"
                                    Write-Host 'System Type            :'"$($HospitalInfo.'System Type')"
                                    Write-Host ''

                                    Clear-Variable -Name HospitalWeb* -ErrorAction Ignore
                                    # Retrieve hospital hours from standard formatted vca site
                                    if ($HospitalInfo.GPURL -and ($($HospitalInfo.'Hospital Number') -notmatch '^[8][0-9]{4}$')) {
                                        try {
                                            $HospitalWeb = Invoke-WebRequest -Uri $HospitalInfo.GPURL -ErrorAction Stop
                                            if ($HospitalWeb) {
                                                $HospitalWebFiltered = $($HospitalWeb.ParsedHtml.body.getElementsByClassName('accordion__item')).innertext
                                                $HospitalWebFiltered
                                                if (-not $HospitalWebFiltered) {
                                                    $HospitalWebFiltered2 = $($HospitalWeb.ParsedHtml.body.getElementsByClassName('hospital-info__middle ')).getElementsByClassName('hospital-info__column col-12 col-md-3 d-flex flex-column')[1].innertext
                                                    $HospitalWebFiltered2
                                                }
                                                <#
                                                if (@(($HospitalWeb.ParsedHtml.body.getElementsByTagName('div') | Where-Object {$PSItem.getAttributeNode('class').Value -eq 'hours'}).textcontent)[0].trim()) {
                                                    Write-Host 'Hours:' -ForegroundColor Cyan
                                                    @(($HospitalWeb.ParsedHtml.body.getElementsByTagName('div') |
                                                        Where-Object {$PSItem.getAttributeNode('class').Value -eq 'hours'}).textcontent)[0].trim() -replace "  ",'' -replace "`n`n",''
                                                }
                                                #>
                                            }
                                        }
                                        catch { Write-Warning $_.Exception.Message }

                                        if (-not $HospitalWeb -or (-not $HospitalWebFiltered -and -not $HospitalWebFiltered2)) {
                                            $HospitalHours = Invoke-RestMethod -Uri "https://uat.vcahospitals.com/api/content/hospital/getUSHospitalHours?HospitalID=$($HospitalInfo.'Hospital Number')"
                                            $HospitalHoursFormatted = $HospitalHours | ForEach-Object {
                                                [pscustomobject]@{
                                                    au                = $_.au
                                                    timezone          = $_.timezone
                                                    monday_open       = $_.monday_open | Get-Date -Format "h:mm tt"
                                                    monday_close      = $_.monday_close | Get-Date -Format "h:mm tt"
                                                    monday_all_day    = $_.monday_all_day
                                                    tuesday_open      = $_.tuesday_open | Get-Date -Format "h:mm tt"
                                                    tuesday_close     = $_.tuesday_close | Get-Date -Format "h:mm tt"
                                                    tuesday_all_day   = $_.tuesday_all_day
                                                    wednesday_open    = $_.wednesday_open | Get-Date -Format "h:mm tt"
                                                    wednesday_close   = $_.wednesday_close | Get-Date -Format "h:mm tt"
                                                    wednesday_all_day = $_.wednesday_all_day
                                                    thursday_open     = $_.thursday_open | Get-Date -Format "h:mm tt"
                                                    thursday_close    = $_.thursday_close | Get-Date -Format "h:mm tt"
                                                    thursday_all_day  = $_.thursday_all_day
                                                    friday_open       = $_.friday_open | Get-Date -Format "h:mm tt"
                                                    friday_close      = $_.friday_close | Get-Date -Format "h:mm tt"
                                                    friday_all_day    = $_.friday_all_day
                                                    saturday_open     = $_.saturday_open | Get-Date -Format "h:mm tt"
                                                    saturday_close    = $_.saturday_close | Get-Date -Format "h:mm tt"
                                                    saturday_all_day  = $_.saturday_all_day
                                                    sunday_open       = $_.sunday_open | Get-Date -Format "h:mm tt"
                                                    sunday_close      = $_.sunday_close | Get-Date -Format "h:mm tt"
                                                    sunday_all_day    = $_.sunday_all_day
                                                    misc_heading      = $_.misc_heading
                                                    misc_hours        = $_.misc_hours -replace '<[^>]+>', ''
                                                }
                                            }
                                            $HospitalHoursFormatted | Out-ListString
                                        }
                                    }
                                    # Canada Hours API
                                    else {
                                        Invoke-RestMethod -Uri "https://uat.vcacanada.com/api/Content/Hospital/GetCAHospitalHours?HospitalID=$($HospitalInfo.'Hospital Number')" | Out-String
                                    }
                                }
                            }
                            Write-Host ''
                        }
                        '14t' {
                            # #14t
                            # generate snow ticket
                            $SiteUsers = Get-VCAHeadcount -AU $ComputerName
                            Invoke-SnowGui -ComputerName $ComputerName -HospitalMaster $HospitalMaster -ImpactedUserList $SiteUsers.Users.EmailAddress
                        }
                        '14u' {
                            # #14u
                            Update-HospitalMaster -EmailCredential $EmailCredential
                            Write-Host "Reloading Hospital Master...`r`n" -ForegroundColor Cyan
                            $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc

                            if (-not $ADCredential) { $ADCredential = Get-ADCreds }
                            if ($ADCredential) {
                                $ClusterSites = Import-Csv -Path "$PSScriptRoot\Private\csv\ClusterSites.csv"
                                Update-ClusterCsv -Clusters $ClusterSites -Credential $ADCredential -Verbose
                            }
                            $ClusterSites = Import-Csv -Path "$PSScriptRoot\Private\csv\ClusterSites.csv"
                        }
                        '15' {
                            # #15
                            # All Hospital Circuits
                            
                            Invoke-UpdateVcaCircuitsCsv
                            Get-VcaHospitalCircuits
                        }
                        '16' {
                            # #16
                            # launch hospital website
                            if (-not $HospitalMaster) {
                                # $HospitalMaster = Import-Csv -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.csv"
                                $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
                            }
                            if ($HospitalMaster) {
                                $HospitalInfo = $HospitalMaster.Where( {
                                        $PSItem."Hospital Number" -eq "$(Convert-VcaAU -AU $ComputerName -Strip)" } )
                                if ($HospitalInfo) {
                                    Write-Output $HospitalInfo | Out-ListString

                                    Write-Host 'Location:'
                                    Write-Host "$($HospitalInfo.'Operating Name') #$($HospitalInfo.'Hospital Number')"
                                    Write-Host "$($HospitalInfo.Address)"
                                    Write-Host "$($HospitalInfo.City), $($HospitalInfo.St) $($HospitalInfo.Zip)"
                                    Write-Host ''
                                    Write-Host 'VCA Site Contact:'
                                    Write-Host "$($HospitalInfo.'Hospital Manager'), $($HospitalInfo.'Hospital Manager Email')"
                                    Write-Host "$($HospitalInfo.Phone)"
                                    Write-Host ''
                                    Write-Host 'Misc. info:'
                                    Write-Host 'Time Zone              :'"$($HospitalInfo.'Time zone')"
                                    Write-Host 'URL                    :'"$($HospitalInfo.GPURL)"
                                    Write-Host 'Back Line              :'"$($HospitalInfo.'Back Line')"
                                    Write-Host 'System Conversion Date :'"$($HospitalInfo.'System Conversion Date')"
                                    Write-Host 'System Type            :'"$($HospitalInfo.'System Type')"

                                    $HospitalURL = $HospitalInfo.GPURL
                                }
                            }
                            if ($HospitalURL) {
                                #'i'm feeling lucky'
                                #Start-Process "https://duckduckgo.com/?q=!ducky+vcahospitals.com+$HospitalName"
                                Start-Process "http://$HospitalURL"
                            }
                            Write-Host ''
                        }
                        '16p' {
                            # #16p
                            # launch hospital website
                            if (-not $HospitalMaster) {
                                $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
                            }
                            if ($HospitalMaster) {
                                $HospitalInfo = $HospitalMaster.Where( {
                                        $PSItem."Hospital Number" -eq "$(Convert-VcaAU -AU $ComputerName -Strip)" } )

                                Write-Host 'Location:'
                                Write-Host "$($HospitalInfo.'Operating Name') #$($HospitalInfo.'Hospital Number')"
                                Write-Host "$($HospitalInfo.Address)"
                                Write-Host "$($HospitalInfo.City), $($HospitalInfo.St) $($HospitalInfo.Zip)"
                                Write-Host ''
                                Write-Host 'VCA Site Contact:'
                                Write-Host "$($HospitalInfo.'Hospital Manager'), $($HospitalInfo.'Hospital Manager Email')"
                                Write-Host "$($HospitalInfo.Phone)"
                                Write-Host ''
                                Write-Host 'Misc. info:'
                                Write-Host 'Time Zone              :'"$($HospitalInfo.'Time zone')"
                                Write-Host 'URL                    :'"$($HospitalInfo.GPURL)"
                                Write-Host 'Back Line              :'"$($HospitalInfo.'Back Line')"
                                Write-Host 'System Conversion Date :'"$($HospitalInfo.'System Conversion Date')"
                                Write-Host 'System Type            :'"$($HospitalInfo.'System Type')"

                                $HospitalURL = $HospitalInfo.GPURL
                            }

                            if ($HospitalURL) {
                                [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="PowerShell Browser: http://$HospitalURL" WindowStartupLocation="CenterScreen">
    <Grid>
        <WebBrowser
            HorizontalAlignment="Left"
            Height="600"
            Margin="10,10,0,0"
            VerticalAlignment="Top"
            Width="1000"
            Name="WebBrowser"
        />
    </Grid>
</Window>
"@
                                $reader = New-Object System.Xml.XmlNodeReader($xaml)
                                $Form = [Windows.Markup.XamlReader]::Load( $reader )
                                $WebBrowser = $Form.FindName('WebBrowser')
                                $WebBrowser.add_Loaded({$this.Navigate("http://$HospitalURL")})

                                $async = $Form.Dispatcher.InvokeAsync( {
                                        $Form.ShowDialog() | Out-Null
                                    })
                                $async.Wait() | Out-Null
                            }
                            Write-Host ''
                        }
                        '17' {
                            # #17
                            # find-hpeilo
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                if ($ComputerName -notmatch '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') {
                                    Clear-Variable -Name ComputerDNS -ErrorAction Ignore
                                    $ComputerDNS = (Resolve-DnsName -Name (Convert-VcaAu -AU $ComputerName -Suffix -gw)).IPAddress
                                }
                                else { $ComputerDNS = $ComputerName }
                                if ($ComputerDNS) {
                                    Write-Host ''
                                    Clear-Variable -Name IloResults -ErrorAction Ignore
                                    Write-Host "PS Command: Find-HPEiLO -Range $(([ipaddress]$ComputerDNS).GetAddressBytes()[0..2] -join '.')`r`n" -ForegroundColor Cyan
                                    $IloResults = Find-HPEiLO -Range (([ipaddress]$ComputerDNS).GetAddressBytes()[0..2] -join '.')
                                    $IloResults | Out-String
                                    $IloResults | foreach-object { Write-Host "https://$($PSItem.IP)" -ForegroundColor Cyan }
                                    Write-Host ''
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '18' {
                            # #18
                            # ipscan
                            Write-Host "Angry IP Scanner Instructions (updated v240702.  Must see a ""Ports [34+]"" column):" -ForegroundColor Cyan
                            Write-Host "1. Navigate to 'Tools'-> 'Preferences...':'Ports' tab and add the following to 'Port selection':" -ForegroundColor Yellow
                            Write-Host "`"80,443,8080,135,137,138,139,445,3389,902,1433,22,23,1270,5723-5724,900,598,161,162,8886,389,636,5900,9100,5989,1720,1730,4649,16992,8443,50000,5985-5986`"" -ForegroundColor Green
                            Write-Host "2. Also consider UnChecking both 'Confirmation' options under 'Tools'-> 'Preferences...':'Display' tab." -ForegroundColor Yellow
                            Write-Host ''
                            Write-Host "Note: To be able to scan non-pinging FUSE devices, Angry IP needs to be altered." -ForegroundColor Green
                            Write-Host "   Check 'Scan dead hosts, which don't reply to pings' under the 'Tools'-> 'Preferences...':'Scanning' tab." -ForegroundColor Green
                            Write-Host ''
                            Write-Host "For Port information visit:" -ForegroundColor Cyan
                            Write-Host "https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.xhtml" -ForegroundColor Gray
                            Write-Host ''
                            Write-Host "Windows:" -ForegroundColor Yellow
                            Write-Host "`t135,137,138,139,445 `t Windows File Transfer (Windows)" -ForegroundColor Gray
                            Write-Host "`t3389 `t`t`t Remote Desktop Server (Windows)" -ForegroundColor Gray
                            Write-Host "`t5985-5986 `t`t WinRM (Windows)" -ForegroundColor Gray
                            Write-Host "ESXi:" -ForegroundColor Yellow
                            Write-Host "`t902 `t`t`t ESXi Host (VMware)" -ForegroundColor Gray
                            Write-Host "ESXi/Linux/iLO/UPS:" -ForegroundColor Yellow
                            Write-Host "`t22 `t`t`t SSH Server (UX-based)" -ForegroundColor Gray
                            Write-Host "Appications:" -ForegroundColor Yellow
                            Write-Host "`t5900 `t`t`t VNC Server" -ForegroundColor Gray
                            Write-Host "`t80,443 - 8080,8443 `t http/https (unsecure/secure) - web services" -ForegroundColor Gray
                            Write-Host "`t1433 `t`t`t SQL Server" -ForegroundColor Gray
                            Write-Host "`t23 `t`t`t Telnet Server (should be disabled/Notify Security Engineering)" -ForegroundColor Gray
                            Write-Host "`t161,162 `t`t SNMP (sender/reciever)" -ForegroundColor Gray
                            Write-Host "`t389,636 `t`t Active Directory (unsecure/secure)" -ForegroundColor Gray
                            Write-Host "`t1270,5723,5724 `t`t Microsoft Operations Manager `"OM`"; OM - Health Service; OM - SDK Service" -ForegroundColor Gray
                            Write-Host "Printers:" -ForegroundColor Yellow
                            Write-Host "`t9100 `t`t`t Standard TCP/IP Printer Port" -ForegroundColor Gray
                            Write-Host "Management Protocols:" -ForegroundColor Yellow
                            Write-Host "`t5989 `t`t`t WBEM (MSAs & etc.)" -ForegroundColor Gray
                            Write-Host "`t16992 `t`t`t Intel AMT (Web Management page of Intel workstations)" -ForegroundColor Gray
                            Write-Host "`t`t`t`t   To aid in finding UTILs (#util) BUT Ping under Network Settings must be" -ForegroundColor Cyan
                            Write-Host "`t`t`t`t   enabled within AMT for when the machine is powered off." -ForegroundColor Cyan
                            Write-Host "Phone systems:" -ForegroundColor Yellow
                            Write-Host "`t1720,1730,4649 `t`t NEC PBX" -ForegroundColor Gray
                            Write-Host "Credit Card Terminals:" -ForegroundColor Yellow
                            Write-Host "`t50000 `t`t`t Ingenico" -ForegroundColor Gray
                            Write-Host "Needs Definition:" -ForegroundColor Yellow
                            Write-Host "`t900,598,8886 `t`t ???, ???, ???" -ForegroundColor Gray
                            Write-Host ''
                            if ($ComputerName -notmatch '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') {
                                Clear-Variable -Name SiteIP, ClusterDNS -ErrorAction Ignore

                                $SiteIP = (Resolve-DnsName -Name $(Convert-VcaAu -AU $ComputerName -Suffix '-gw')).IPAddress
                                if ($Cluster) {
                                    $ClusterDNS = (Resolve-DnsName -Name $ComputerName).IPAddress
                                    $SiteIP = @($SiteIP, $ClusterDNS)
                                }
                            }
                            else { $SiteIP = $ComputerName }

                            if (Get-Command -Name java -ErrorAction Ignore) {
                                if ((Get-Command -Name java).Version.Major -lt 11) {
                                    Write-Host "`r`nJava $((Get-Command -Name java).Version.ToString()) detected, please install Microsoft's OpenJDK package." -ForegroundColor Cyan
                                    Write-Host "Installer: `"\\vcaantech.com\folders\Apps\install_media\Microsoft\OpenJDK\17.06 LTS`"" -ForegroundColor Cyan
                                    Write-Host "`r`n* Please close out all portal sessions after OpenJDK is install and relaunch if needed.`r`n" -ForegroundColor Cyan
                                    break
                                }
                            }
                            else {
                                Write-Host "`r`nJava not detected, please install Microsoft's OpenJDK." -ForegroundColor Cyan
                                Write-Host "Installer: `"\\vcaantech.com\folders\Apps\install_media\Microsoft\OpenJDK\17.06 LTS`"" -ForegroundColor Cyan
                                Write-Host "`r`n* Please close out all portal sessions after OpenJDK is install and relaunch if needed.`r`n" -ForegroundColor Cyan
                                break
                            }
                            $SiteIP | foreach-object {
                                Clear-Variable -Name Subnet -ErrorAction Ignore
                                $Subnet = (([ipaddress]$PSItem).GetAddressBytes()[0..2] -join '.')
                                $StartIP = $Subnet + '.1'
                                $EndIP = $Subnet + '.254'

                                Start-Process -FilePath 'ipscan-win64-3.9.0.exe' -ArgumentList "-s -f:range $StartIP $EndIP" -WorkingDirectory "$PSScriptRoot\Private\bin"
                            }
                        }
                        '19?' {
                            # #19?
                            Write-Host "RDC Launch shortcuts" -ForegroundColor Cyan
                            Write-Host " 19n  -ns and -nslb" -ForegroundColor Cyan
                            Write-Host " 19d  -db" -ForegroundColor Cyan
                            Write-Host " 19fs -fs" -ForegroundColor Cyan
                            Write-Host " 19s  -smpacs" -ForegroundColor Cyan
                            Write-Host " 19cs -cs" -ForegroundColor Cyan
                            Write-Host " 19u  -util" -ForegroundColor Cyan
                            Write-Host "Web Launch shortcuts" -ForegroundColor Cyan
                            Write-Host " 19sw -smpacs login & scom values" -ForegroundColor Cyan
                            Write-Host "19sws -smpacs status" -ForegroundColor Cyan
                        }
                        '19' {
                            # #19
                            # rdc
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers, SiteAU -ErrorAction Ignore
                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                    Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                    Out-GridView -Title "#19 Select Remote Desktop Server(s) to launch - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null
                                if (@($SiteServers).count -ge 1) { Write-Host "`r`n Launching Remote Desktop Client:`r`n" -ForegroundColor Cyan }
                                $SiteServers | foreach-object {
                                    Write-Host "  > mstsc.exe /v:$($PSItem.Name) /admin" -ForegroundColor Cyan
                                    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$($PSItem.Name) /admin"
                                }
                                if (@($SiteServers).count -ge 1) { Write-Host "`r" -ForegroundColor Cyan }
                            }
                            else {
                                if ((-not $Cluster) -and $NetServices) {
                                    Write-Host "`r`n Launching Remote Desktop Client:`r`n" -ForegroundColor Cyan
                                    Write-Host "  > mstsc.exe /v:$NetServices /admin" -ForegroundColor Cyan
                                    Start-Process -FilePath "> mstsc.exe" -ArgumentList "/v:$NetServices"
                                    Write-Host "`r" -ForegroundColor Cyan
                                }
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '19cs' {
                            # #19cs
                            Clear-Variable -Name Cornerstone -ErrorAction Ignore
                            $Cornerstone = (Get-ADComputer -Filter "Name -like '$(Convert-VcaAU -AU $ComputerName -Suffix -cs)*'").Name
                            $Cornerstone | ForEach-Object {
                                if ($PSItem | Get-PingStatus) {
                                    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$PSItem"
                                }
                            }
                        }
                        '19d' {
                            # #19d
                            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$(Convert-VcaAU -AU $ComputerName -Suffix '-db')"
                        }
                        '19fs' {
                            # #19fs
                            Clear-Variable -Name FileServer -ErrorAction Ignore
                            $FileServer = (Get-ADComputer -Filter "Name -like '$(Convert-VcaAU -AU $ComputerName -Suffix -fs)*' -and OperatingSystem -notlike 'Windows Server 2008*'").Name
                            $FileServer | ForEach-Object {
                                if ($PSItem | Get-PingStatus) {
                                    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$PSItem"
                                }
                            }
                        }
                        '19n' {
                            # #19n
                            if (-not $Cluster) {
                                Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$(Convert-VcaAU -AU $ComputerName -Suffix '-ns')"
                            }
                            elseif ($Cluster) {
                                Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$(Convert-VcaAU -AU $ComputerName -Suffix '-nslb')"
                            }
                        }
                        '19s' {
                            # #19s
                            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$(Convert-VcaAU -AU $ComputerName -Suffix '-smpacs')"
                        }
                        '19sw' {
                            # #19sw
                            Clear-Variable -Name SmpacsDiskResults -ErrorAction Ignore
                            $SiteSmpacs = Convert-VcaAU -AU $ComputerName -Suffix '-smpacs'
                            $SmpacsDiskResults = Get-DiskUsage -ComputerName $SiteSmpacs
                            ($SmpacsDiskResults | Select-Object -Property ComputerName, Name, FileSystem, FreeGB, 'FreeSpace(%)', CapacityGB,
                                @{n='SCOMCap50GB(%)';e={[decimal]('{0:N0}' -f (100 - (50 / $_.CapacityGB) * 100))}} | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''

                            Start-Process "http://$SiteSmpacs/SmartPACS/"
                        }
                        '19sws' {
                            # #19sws
                            Start-Process "http://$(Convert-VcaAU -AU $ComputerName -Suffix '-smpacs')/SmartPACS/serverstatus"
                        }
                        '19u' {
                            # #19u
                            Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$(Convert-VcaAU -AU $ComputerName -Suffix '-util')"
                        }
                        '20' {
                            # #20
                            # active directory computers
                            Get-VcaADComputers -AU $ComputerName | Out-TableString -Wrap
                        }
                        '21' {
                            # #21
                            # launch orion
                            Start-Process "https://phvcaorionp01/Orion/NetPerfMon/Resources/NodeSearchResults.aspx?Property=Caption&SearchText=$(Convert-VcaAu $ComputerName -Strip)&ResourceID=9"
                            Write-Host ''
                        }
                        '22' {
                            # #22
                            # site pictures
                            Get-ChildItem -Path "\\vcaantech.com\folders\data2\Corp\Information Technology\Site_Pictures\*\$(Convert-VcaAU -AU $ComputerName -Strip)-*" -Directory | Invoke-Item
                        }
                        '22w' {
                            # #22w
                            # site pictures web
                            Start-Process "https://vca365.sharepoint.com/sites/CentralBark/SupportOffice/IT/LocationDetail/Forms/AllItems.aspx?viewid=4a2ae99d%2D2d90%2D46f3%2D81fe%2D007738d7e509"
                        }
                        '23' {
                            # #23
                            # display dns cname and a records
                            if ((Get-Module -Name DnsServer -ListAvailable) -or (Get-Module -Name DnsServer)) {
                                if (-not $DNSServer) {
                                    # Check if already domain admin
                                    if ((whoami /groups) -match 'Domain Admins') {
                                        $DNSServer = New-CimSession -ComputerName ladcp01
                                    }
                                    elseif (-not $ADCredential) {
                                        $ADCredential = Get-ADCreds
                                    }

                                    Write-Host "Attempting to Connect with: " $ADCredential
                                    if ($ADCredential) {
                                        # Save authenticated session
                                        $DNSServer = New-CimSession -ComputerName ladcp01 -Credential $ADCredential -SessionOption (New-CimSessionOption -Protocol Dcom)
                                    }
                                }
                                if ($DNSServer) {
                                    $DNSCName = Get-DnsServerResourceRecord -CimSession $DNSServer -ZoneName 'vcaantech.com' -RRType CName |
                                        Where-Object HostName -like "$(Convert-VcaAU -AU $ComputerName -Suffix '')-*"
                                    $DNSARecord = Get-DnsServerResourceRecord -CimSession $DNSServer -ZoneName 'vcaantech.com' -RRType A |
                                        Where-Object HostName -like "$(Convert-VcaAU -AU $ComputerName -Suffix '')-*"

                                    (@($DNSCName) + @($DNSARecord)) | Out-GridView -Title "#23 Select entries to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutVariable DNSRecords -OutputMode Multiple |
                                        Format-Table -AutoSize | Out-String
                                }
                                if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                    New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                }
                            }
                            else {
                                Write-Warning 'DnsServer module not found.'
                                Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '24' {
                            # #24
                            # Launch DHCP MMC
                            if (Get-Module -Name ActiveDirectory) {
                                if (-not $ADCredential) {
                                    $ADCredential = Get-ADCreds
                                }
                                if ($ADCredential) {
                                    Clear-Variable -Name SiteDC -ErrorAction Ignore

                                    [System.Collections.ArrayList]$DhcpServers = @()
                                    $SiteDC = Get-ADComputer -Filter "Name -like '$(Convert-VcaAu -AU $ComputerName -Suffix '-dc')*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address
                                    $PhoenixDC = Get-ADComputer -Filter "Name -like 'PHHOSPDHCP*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address

                                    (@($SiteDC) + @($PhoenixDC)) | ForEach-Object {
                                        $DhcpServers.Add(( $PSItem |
                                        Select-Object -Property @{n='Name';e={$PSItem.Name}},
                                                                @{n='ADIPv4Address';e={$PSItem.IPv4Address}},
                                                                @{n='CanonicalName';e={$PSItem.CanonicalName}},
                                                                @{n='Status';e={$PSItem.Name | Get-PingStatus}}
                                                                )) | Out-Null
                                    }
                                    $DhcpServers | Out-GridView -Title "#24 Select DHCP MMC(s) to launch - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable DhcpServerSelection | Out-Null
                                    if ($DhcpServerSelection) {
                                        $DhcpServerSelection.Name | ForEach-Object {
                                            Start-Process -FilePath cmd.exe -Credential $ADCredential -ArgumentList ("/c dhcpmgmt.msc /ComputerName $PSItem") -WorkingDirectory "$env:SystemRoot\System32" -WindowStyle Hidden
                                        }
                                    }
                                }
                                if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                    New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                }
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '24a' {
                            # #24a
                            Clear-Variable -Name SiteDC, SiteIP, DhcpServerSelection, DhcpScopeSelection -ErrorAction Ignore

                            [System.Collections.ArrayList]$DhcpServers = @()
                            $SiteIP = (Resolve-DnsName -Name $(Convert-VcaAu -AU $ComputerName -Suffix '-gw')).IPAddress
                            $SiteDC = Get-ADComputer -Filter "Name -like '$(Convert-VcaAu -AU $ComputerName -Suffix '-dc')*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address
                            $PhoenixDC = Get-ADComputer -Filter "Name -like 'PHHOSPDHCP*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address

                            #(@($SiteDC) + @($PhoenixDC)[0]) | ForEach-Object {
                            @($PhoenixDC)[0] | ForEach-Object {

                                $DhcpServers.Add(( $PSItem |
                                Select-Object -Property @{n='Name';e={$PSItem.Name}},
                                                        @{n='ADIPv4Address';e={$PSItem.IPv4Address}},
                                                        @{n='CanonicalName';e={$PSItem.CanonicalName}},
                                                        @{n='Status';e={$PSItem.Name | Get-PingStatus}}
                                                        )) | Out-Null
                            }
                            if ($DhcpServers) {
                                if (-not $ADCredential) { $ADCredential = Get-ADCreds }
                                $SitePrefix = Convert-VcaAu -AU $ComputerName -Strip
                                # centralized dhcp server selected
                                try {
                                    #Write-Host " > Invoke-Command -ComputerName $($DhcpServers.Name -join ', ') `{ Get-DhcpServerv4Scope | Where-Object `{ `$_.Name -match `"^h?0?$SitePrefix`" `} `}" -ForegroundColor Cyan
                                    #$DhcpScopes = Invoke-Command -ComputerName $DhcpServers.Name { Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*$using:SitePrefix*" } } -Credential $ADCredential -ErrorAction Stop |
                                    #$DhcpScopes = Invoke-Command -ComputerName $DhcpServers.Name { Get-DhcpServerv4Scope | Where-Object { $_.Name -match "^h?0?$using:SitePrefix" } } -Credential $ADCredential -ErrorAction Stop |
                                    #Write-Host " > Invoke-Command -ComputerName DhcpServers.Name `{ Get-DhcpServerv4Scope | Where-Object `{ `$_.Name -match `"^h?0?$SitePrefix`" `} `}" -ForegroundColor Cyan
                                    #$DhcpScopes = Invoke-Command -ComputerName $DhcpServers.Name { Get-DhcpServerv4Scope | Where-Object { $_.Name -match "^$SitePrefix[A-Za-z]?$|^$SitePrefix[A-Za-z]?? |h$SitePrefix ?" } } -Credential $ADCredential -ErrorAction Stop
                                    Write-Host " > Invoke-Command -ComputerName $($DhcpServers.Name -join ', ') `{ Get-DhcpServerv4Scope | Where-Object `{ `$_.Name -like `"*AU$SitePrefix-*`" `} `}" -ForegroundColor Cyan
                                    $DhcpScopes = Invoke-Command -ComputerName $DhcpServers.Name { Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$using:SitePrefix-*" } } -Credential $ADCredential -ErrorAction Stop |
                                        Select-Object -Property PSComputerName, ScopeId, SubnetMask, Name, State, StartRange, EndRange, LeaseDuration

                                    if ($DhcpScopes.ScopeId) {
                                        $DhcpScopes | ForEach-Object {
                                            $DhcpScopes_Item = $PSItem
                                            Write-Host " > Get-DhcpServerv4OptionValue -ComputerName $($DhcpScopes_Item.PSComputerName) -ScopeId $($PSItem.ScopeId) #Scope Name: $($PSItem.Name)" -ForegroundColor Cyan

                                            Invoke-Command -ComputerName $DhcpScopes_Item.PSComputerName { Get-DhcpServerv4OptionValue -ScopeId $using:PSItem.ScopeId } -Credential $ADCredential -ErrorAction Stop |
                                                Select-Object -Property @{n='ScopeName';e={$DhcpScopes_Item.Name}}, @{n='ScopeId';e={$DhcpScopes_Item.ScopeId}}, OptionId, Name, Value, PolicyName, Type, VendorClass, UserClass, PSComputerName
                                        } | Sort-Object -Property PSComputerName, ScopeName, ScopeId | Out-GridView -PassThru -Title "#24a AU$SitePrefix - DhcpServerv4Lease - Select entries to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" |
                                                Format-Table -AutoSize -Wrap | Out-String
                                    }
                                }
                                catch {
                                    Write-Warning $_.Exception.Message
                                }
                            }
                            if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                            }
                        }
                        '24b' {
                            # #24b
                            # mac list source: https://gitlab.com/wireshark/wireshark/-/raw/master/manuf

                            Clear-Variable -Name SiteDC, DhcpServerSelection -ErrorAction Ignore

                            [System.Collections.ArrayList]$DhcpServers = @()
                            $SiteDC = Get-ADComputer -Filter "Name -like '$(Convert-VcaAu -AU $ComputerName -Suffix '-dc')*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address
                            $PhoenixDC = Get-ADComputer -Filter "Name -like 'PHHOSPDHCP*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties CanonicalName, IPv4Address

                            (@($PhoenixDC) + @($SiteDC)) | ForEach-Object {
                                $DhcpServers.Add(( $PSItem |
                                Select-Object -Property @{n='Name';e={$PSItem.Name}},
                                                        @{n='ADIPv4Address';e={$PSItem.IPv4Address}},
                                                        @{n='CanonicalName';e={$PSItem.CanonicalName}},
                                                        @{n='Status';e={$PSItem.Name | Get-PingStatus}}
                                                        )) | Out-Null
                            }
                            $DhcpServers | Out-GridView -Title "#24b Select DHCP server to query - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable DhcpServerSelection | Out-Null
                            if ($DhcpServerSelection) {
                                if (-not $ADCredential) { $ADCredential = Get-ADCreds }
                                try {
                                    Invoke-DhcpPrompt -ComputerName $ComputerName -DhcpServer $DhcpServerSelection.Name -Credential $ADCredential -ErrorAction Stop

                                    if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                        New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                    }
                                }
                                catch {
                                    Write-Warning $_.Error.Exception.Message
                                }
                            }
                        }
                        '24c' {
                            # #24c
                            $SiteIP = (Resolve-DnsName -Name $(Convert-VcaAu -AU $ComputerName -Suffix '-gw')).IPAddress
                            Write-Host "Get-DhcpServerv4Reservation -ComputerName phhospdhcp1 -ScopeId $(([ipaddress]$SiteIP).GetAddressBytes()[0..2] -join '.').0" -ForegroundColor Cyan
                            Get-DhcpServerv4Reservation -ComputerName phhospdhcp1 -ScopeId "$(([ipaddress]$SiteIP).GetAddressBytes()[0..2] -join '.').0" | Out-GridView -PassThru -Title "#24c DhcpServerv4reservation - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" |
                                Sort-Object -Property IPAddress | Format-Table -AutoSize -Wrap | Out-String
                        }
                        '25' {
                            # #25
                            # Restart windows services
                            Clear-Variable -Name SiteServices, SiteServer -ErrorAction Ignore
                            $SiteServer = Select-VcaSite -AU $Computername -Title "#25 Select Remote Desktop Server to query services - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single

                            if ($SiteServer) {
                                if (-not $ADCredential) { $script:ADCredential = Get-ADCreds }
                                Invoke-command -ComputerName $SiteServer.Name -Credential $ADCredential -ScriptBlock {
                                    Get-Service | Select-Object -Property Name, DisplayName, Status, MachineName | Sort-Object -Property Name
                                 } | Out-GridView -Title "#25 [$($SiteServer.Name)] - Select Service to restart - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable SiteServices | Out-Null
                                if ($SiteServices) {
                                    $SiteServices | Out-String
                                    if ((Read-Choice -Title "#25 Restart [$($SiteServer.Name) - $($SiteServices.Name)] - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 0) -eq 1) {
                                        #Get-Service -ComputerName $SiteServer.Name -Name $SiteServices.Name | Restart-Service -Force -Verbose
                                        #Get-Service -ComputerName $SiteServer.Name -Name $SiteServices.Name -Verbose | Format-Table -AutoSize | Out-String
                                        Invoke-command -ComputerName $SiteServer.Name -Credential $ADCredential -ScriptBlock {
                                            Get-Service -Name $using:SiteServices.Name | Restart-Service -Force -Verbose
                                            Get-Service -Name $using:SiteServices.Name -Verbose | Format-Table -AutoSize | Out-String
                                        }
                                    }
                                }
                                if ($ADCredential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
                                    New-StoredCredential -Credentials $ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                }
                            }
                        }
                        '25sql' {
                            # #25sql
                            # Restart MSSQL services on DB
                            $DBComputer = "$(Convert-VcaAU -AU $ComputerName -Suffix '-db')"
                            Invoke-Command -ComputerName $DBComputer {
                                Get-Service -Name "MSSQLServer" | Restart-Service -Force -Confirm:$true
                            }
                        }
                        '26' {
                            # #26
                            # Links/Portals
                            Clear-Variable -Name LinksPortalRemote -ErrorAction Ignore
                            try {
                                $LinksPortalCsv = Import-Csv -Path "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\WebPortals.csv" -ErrorAction Stop
                                $LinksPortalRemote = $true
                            }
                            catch {
                                Write-Warning $_.Exception.Message

                                Write-Host "`r`nReading from local copy" -ForegroundColor Cyan
                                $LinksPortalCsv = Import-Csv -Path "$PSScriptRoot\private\csv\WebPortals.csv"
                                $LinksPortalRemote = $false
                            }
                            $ContactsInfo = [PSCustomObject]@{
                                Site     = "Ops Portal Tool's Contacts (#98)"
                                URL      = $(if($LinksPortalRemote) {'\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\Contacts.csv'} else {"$PSScriptRoot\private\csv\Contacts.csv"})
                                ID       = ''
                                Password = ''
                            }
                            $WebPortalsInfo = [PSCustomObject]@{
                                Site     = "Ops Portal Tool's Links/Portals (#26/This document)"
                                URL      = $(if($LinksPortalRemote) {'\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\WebPortals.csv'} else {"$PSScriptRoot\private\csv\WebPortals.csv"})
                                ID       = ''
                                Password = ''
                            }
                            $LinksPortalCsv + $ContactsInfo + $WebPortalsInfo |
                                Out-GridView -Title "#26 Select link(s) to launch $(if($LinksPortalRemote) {'- Remote: \\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\WebPortals.csv'} else {"- Local: $PSScriptRoot\private\csv\WebPortals.csv"}) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple |
                                    ForEach-Object {
                                        $PSItem
                                        Start-Process $PSItem.URL
                                    } | Out-String
                        }
                        '26l' {
                            # #26L
                            # launch csvs directory
                            Invoke-Item -Path "$PSSCriptRoot\Private\csv"
                        }
                        '27' {
                            # #27
                            # Query ilo status
                            if (-not $IloCredential ) { $IloCredential = Get-StoredCredential -Target vcahospilo }
                            Get-IloStatus -NetServices $NetServices -Cluster $Cluster -ComputerName $ComputerName -Credential $IloCredential
                        }
                        '27d' {
                            # #27d
                            # Query disks via ilo
                            if (-not $IloCredential ) { $IloCredential = Get-StoredCredential -Target vcahospilo }
                            if ((-not $Cluster) -and $NetServices) {
                                $ServerIlo = "$ComputerName-ilo"
                            }
                            else {
                                $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                            }
                            Get-VCAHPEDriveFirmwareInfo -ComputerName $ServerIlo -Credential $IloCredential | Out-GridView -Title "#27d Select entries to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -PassThru | Out-String
                            Clear-Variable -Name ServerIlo -WarningAction Ignore
                        }
                        '27p' {
                            # #27p
                            if (-not $IloCredential ) { $IloCredential = Get-StoredCredential -Target vcahospilo }
                            Get-IloBIOSPower -NetServices $NetServices -Cluster $Cluster -ComputerName $ComputerName -Credential $IloCredential
                        }
                        '28' {
                            # #28
                            # Query esxi host
                            if (Get-Module -Name VMware.VimAutomation.Core -ListAvailable) {
                                # Load credentials
                                if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                                if ($EsxiCredential) {
                                    Clear-Variable -Name ESXiInfo, HostVMs -ErrorAction Ignore
                                    try {
                                        $VIServer = Connect-VIServer -Server $ComputerName -Credential $EsxiCredential -WarningAction SilentlyContinue -ErrorAction Stop
                                        if ($VIServer.IsConnected) {
                                            # VM Host
                                            $ESXiInfo = Get-View -Server $VIServer -ViewType HostSystem -Property Name, Runtime.BootTime, Hardware.SystemInfo, Config.PowerSystemInfo.CurrentPolicy
                                            ($ESXiInfo | Select-Object -Property Name,
                                                @{n = 'Model'; e = { $_.Hardware.SystemInfo.Model } },
                                                @{n = 'SerialNumber'; e = { ($_.Hardware.SystemInfo.OtherIdentifyingInfo | Where-Object { $_.IdentifierType.Label -like 'Service tag' }).IdentifierValue } },
                                                @{n = 'BootTime'; e = { $_.Runtime.BootTime } },
                                                @{n = 'ESXiPowerPolicy';e={$_.Config.PowerSystemInfo.CurrentPolicy.ShortName}} | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n\r\n', ''

                                            (Get-VMHost -Server $VIServer | Select-Object -Property Name, ConnectionState, PowerState, NumCpu,
                                                @{n='MemoryUsageGB';e={[math]::Round($_.MemoryUsageGB,2)}},
                                                @{n='MemoryTotalGB';e={[math]::Round($_.MemoryTotalGB,2)}},
                                                Version | Format-Table -AutoSize -Wrap | Out-String) -replace '\r\n\r\n\r\n', ''
                                            # Datastore
                                            (Get-Datastore -Server $VIServer -Refresh | Select-Object -Property @{n = 'DatastoreName'; e = { $_.Name } },
                                                @{n='FreeSpaceGB';e={[math]::Round($_.FreeSpaceGB,2)}}, @{n='CapacityGB';e={[math]::Round($_.CapacityGB,2)}}, @{n='FreeSpacePercent';e={[math]::Round(($_.FreeSpaceGB / $_.CapacityGB) * 100,2)}} | Out-String) -replace '\r\n\r\n\r\n', ''
                                            # VM
                                            $HostVMs = Get-VM -Server $VIServer
                                            ($HostVMs | Select-Object -Property Name, @{n='UsedSpaceGB';e={[math]::Round($_.UsedSpaceGB,2)}},
                                                @{n='ProvisionedSpaceGB';e={[math]::Round($_.ProvisionedSpaceGB,2)}},
                                                @{n='DatastoreName';e={($_ | Get-Datastore).Name}},
                                                @{n='FreeSpaceGB';e={[math]::Round(($_ | Get-Datastore).FreeSpaceGB,2)}},
                                                @{n='CapacityGB';e={[math]::Round(($_ | Get-Datastore).CapacityGB,2)}} |
                                                Sort-Object -Property DatastoreName | Format-Table -AutoSize -Wrap | Out-String) -replace '\r\n\r\n\r\n', ''
                                            ($HostVMs | Select-Object -Property Name, PowerState, NumCpu, CoresPerSocket, MemoryGB,
                                                @{n='IPAddress';e={$_.Guest.IPAddress}}, @{n='VMWareToolsState';e={$_.Guest.State}} |
                                                Sort-Object -Property MemoryGB -Descending | Format-Table -AutoSize -Wrap | Out-String) -replace '\r\n\r\n\r\n', ''

                                            Write-Host ''

                                            # Save credentials
                                            Set-EsxiCredential -Credential $EsxiCredential
                                        }
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                }
                            }
                            else {
                                Write-Warning 'VMware.VimAutomation.Core module not found.'
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name VMware.PowerCLI"
                            }
                        }
                        '28l' {
                            # #28L
                            # check for disk latency on host server
                            # Load credentials
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential) {
                                try {
                                    $VIServer = Connect-VIServer -Server $ComputerName -Credential $EsxiCredential -WarningAction SilentlyContinue -ErrorAction Stop
                                    if ($VIServer.IsConnected) {
                                        Write-Host "Checking host server for storage latency.  Frequent values over 20ms may indicate issues with storage:"-ForegroundColor Magenta
                                        Get-Stat -Server $VIServer -Realtime -MaxSamples 10 -Stat datastore.totalReadLatency.average ,datastore.totalWriteLatency.average, datastore.maxTotalLatency.latest |
                                            Sort-Object -Property Instance, MetricId,TimeStamp | Select-Object -Property Entity,Timestamp, MetricId, Value, Unit, Instance | Out-TableString
                                    }
                                }
                                catch {
                                    Write-Warning $_.Exception.Message
                                }
                            }
                        }
                        '28n' {
                            # #28n
                            # esxi networking
                            # Load credentials
                            if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                            if ($EsxiCredential) {
                                $VIServer = Connect-VIServer -Server $ComputerName -Credential $EsxiCredential -WarningAction SilentlyContinue
                                if ($VIServer.IsConnected) {
                                    Get-VMHostNetworkAdapter -Server $VIServer | Select-Object -Property Name, PortGroupName, MTU, BitRatePerSec, FullDuplex, Mac, DhcpEnabled, IP, SubnetMask | Out-TableString
                                    Get-VMHostNetworkAdapter -Server $VIServer | Select-Object -Property Name,
                                        @{n = 'AutoNegotiate';e ={if ($_.ExtensionData.Spec.LinkSpeed) { $false } else { $true }}},
                                        @{n = 'LinkState';e={if ($_.ExtensionData.LinkSpeed) { 'Up' } else { 'Down' }}} | Out-TableString
                                    Get-VirtualSwitch -Server $VIServer | Select-Object -Property Name, Nic, Mtu | Out-TableString

                                    $LoadBalancingPolicy = {
                                        "$($_.LoadBalancingPolicy)$(
                                            if ($_.LoadBalancingPolicy -eq 'LoadBalanceSrcId') {' (Originating vPort)'}
                                            elseif ($_.LoadBalancingPolicy -eq 'LoadBalanceIP') {' (IP Hash)'}
                                        )"
                                    }
                                    Get-VirtualSwitch -Server $VIServer | Get-NicTeamingPolicy | Select-Object -Property VirtualSwitch, @{n='LoadBalancingPolicy';e=$LoadBalancingPolicy} | Out-TableString
                                    Get-VirtualPortGroup -Server $VIServer | Get-NicTeamingPolicy | Select-Object -Property VirtualPortGroup, @{n='LoadBalancingPolicy';e=$LoadBalancingPolicy} | Out-TableString
                                    Get-VirtualPortGroup -Server $VIServer | Select-Object -Property Name, VLanID | Format-Table -AutoSize | Out-TableString

                                    Write-Host ''

                                    # Save credentials
                                    Set-EsxiCredential -Credential $EsxiCredential
                                }
                            }
                        }
                        '29' {
                            # #29
                            # VMWare Remote Console
                            if (Get-Module -Name VMware.VimAutomation.Core -ListAvailable) {
                                # Load Creds
                                if (-not $EsxiCredential) { $EsxiCredential = Get-EsxiCredential }
                                if ($EsxiCredential) {

                                    $VIServer = Connect-VIServer -Server $ComputerName -Credential $EsxiCredential -WarningAction SilentlyContinue
                                    if ($VIServer.IsConnected) {
                                        Clear-Variable -Name SiteServer -ErrorAction Ignore

                                        Get-VM -Server $VIServer | Select-Object -Property Name, PowerState, @{n = 'IPAddress'; e = { $_.Guest.IpAddress } },
                                        @{n = 'OSFullName'; e = { $_.Guest.OSFullName } }, Id |
                                        Out-GridView -Title "#29 [$ComputerName] Launch VMWare Remote Console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable SiteServer | Out-Null

                                        if ($SiteServer) {
                                            if (-not $VMRCInstall) {
                                                $VMRCInstall = (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                                                    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*').Where( { $_.DisplayName -like 'VMware Remote Console*' } )
                                            }
                                            if ($VMRCInstall) {
                                                Start-Process "vmrc://$($EsxiCredential.UserName)@$ComputerName/?moid=$($SiteServer.Id -replace 'VirtualMachine-','')"
                                            }
                                            else {
                                                Start-Process "$PSScriptRoot\Private\lib\VMware.VimAutomation.Core\11.3.0.13964826\net45\VMware Remote Console\vmrc.exe" -ArgumentList "vmrc://$($EsxiCredential.UserName)@$ComputerName/?moid=$($SiteServer.Id -replace 'VirtualMachine-','')"

                                            }
                                        }
                                        Set-EsxiCredential -Credential $EsxiCredential
                                    }
                                }
                            }
                            else {
                                Write-Warning 'VMware.VimAutomation.Core module not found.'
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name VMware.PowerCLI"
                            }
                        }
                        '30' {
                            # #30
                            # check critical services
                            if (Get-Module -Name ActiveDirectory) {
                                if (-not $NetServices) { break }
                                Clear-Variable -Name DBServices, NSServices, Services, SiteAU -ErrorAction Ignore

                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                $SiteDB = Get-ADComputer -Filter "Name -like '$SiteAU-db*' -and Name -notlike '*CNF:*' -and Enabled -eq '$true'" |
                                    Select-Object -ExpandProperty Name

                                if (-not $Cluster -and -not $VMCSDDCCluster) {
                                    # Single server site
                                    try {
                                        Invoke-Command -ComputerName $SiteDB {
                                            $Service = Get-Service -Name MSSQLSERVER, SQLSERVERAGENT, *Report*Serv* #ReportServer, SQLServerReportingServices
                                            $Service | Select-Object -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $Service | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose
                                                $Service | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    try {
                                        Invoke-Command -ComputerName $NetServices {
                                            $Service = Get-Service -Name SparkyAntechWinService, SparkyEmailService, tssdis, Spooler
                                            $Service | Select-Object -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            Get-Process -Name tssdis | Select-Object -Property ProcessName, StartTime, @{n='OSLastBootupTime';e={(Get-CimInstance -ClassName Win32_OperatingSystem).LastBootupTime}},
                                                @{n='DateTime(Local)';e={"{0:M/dd/yyyy h:mm tt} $([Regex]::Replace([System.TimeZoneInfo]::Local.Id, '([A-Z])\w+\s*', '$1'))" -f (Get-Date)}},
                                                @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $IISPOOL = Get-IISAppPool -Name "WoofwareService", "WoofwareAPI"
                                            $IISPOOL | Select-Object -Property Name, State, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $Service | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose -Confirm:$true
                                                $Service | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} -AutoSize | Out-String
                                            }
                                            $PoolRestart = $IISPOOL | Where-Object State -eq Stopped
                                            if ($PoolRestart) {
                                                Start-WebAppPool -Name WoofwareAPI -Verbose ; Start-WebAppPool WoofwareService -Verbose
                                                $IISPOOL | GET-IISAppPool | Select-Object -Property Name, State, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                }
                                else {
                                    # Cluster
                                    Clear-Variable -Name FSServices -ErrorAction Ignore

                                    $SiteNS = Convert-VcaAu -AU $ComputerName -Suffix '-api'
                                    $SiteNS = Resolve-DnsName $SiteNS | Where-Object Name -match '-ns' | Select-Object -ExpandProperty Name
                                    $SiteNS = $SiteNS -replace '.vcaantech.com', ''

                                    $SiteFS = Get-ADComputer -Filter "Name -like '$SiteAU-fs*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" |
                                        Select-Object -ExpandProperty Name

                                    try {
                                        Invoke-Command -ComputerName $SiteDB {
                                            $Service = Get-Service -Name MSSQLSERVER, SQLSERVERAGENT, *Report*Serv* #ReportServer, SQLServerReportingServices
                                            $Service | Select-Object -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $Service | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose -Confirm:$true
                                                $Service | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    try {
                                        Invoke-Command -ComputerName $SiteNS {
                                            $Service = Get-Service -Name SparkyAntechWinService, SparkyEmailService, Spooler
                                            $Service | Select-Object -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String
                                            $IISPOOL = Get-IISAppPool -Name "WoofwareService", "WoofwareAPI"
                                            $IISPOOL | Select-Object -Property Name, State, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $Service | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose
                                                $Service | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} -AutoSize | Out-String
                                            }
                                            $PoolRestart = $IISPOOL | Where-Object State -eq Stopped
                                            if ($PoolRestart) {
                                                Start-WebAppPool -Name WoofwareAPI -Verbose ; Start-WebAppPool WoofwareService -Verbose
                                                $IISPOOL | GET-IISAppPool | Select-Object -Property Name, State, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    try {
                                        Invoke-Command -ComputerName $SiteFS {
                                            $Service = Get-Service -Name tssdis
                                            $Service | Select-Object -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String


                                            Get-Process -Name tssdis | Select-Object -Property ProcessName, StartTime, @{n='OSLastBootupTime';e={(Get-CimInstance -ClassName Win32_OperatingSystem).LastBootupTime}},
                                                @{n='DateTime(Local)';e={"{0:M/dd/yyyy h:mm tt} $([Regex]::Replace([System.TimeZoneInfo]::Local.Id, '([A-Z])\w+\s*', '$1'))" -f (Get-Date)}},
                                                @{n='ComputerName';e={$env:COMPUTERNAME}} | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $Service | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose
                                                $Service | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n='ComputerName';e={$env:COMPUTERNAME}} -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                }
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                            ''
                        }
                        '30s' {
                            # #30s
                            # check system management services
                            if (Get-Module -Name ActiveDirectory -ListAvailable) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#30s Select Remote Desktop Server(s) to check system management services - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                                $SiteServers | foreach-object {
                                    try {
                                        Write-Host  "`r`n$($PSItem.name)"  -ForegroundColor Cyan
                                        Invoke-Command -ComputerName $PSItem.name {
                                            $ServiceI = Get-Service -Name healthservice, Druva-EnterpriseWorkloadsSVC, PhoenixCPHService
                                            $ServiceI | Select-Object -Property Name, DisplayName, Status, @{n = 'ComputerName'; e = { $env:COMPUTERNAME } } | Format-Table -AutoSize | Out-String

                                            $ServiceRestart = $ServiceI | Where-Object Status -eq Stopped
                                            if ($ServiceRestart) {
                                                $ServiceRestart | Start-Service -Verbose -Confirm:$true
                                                $ServiceI | Get-Service | Format-Table -Property Name, DisplayName, Status, @{n = 'ComputerName'; e = { $env:COMPUTERNAME } } -AutoSize | Out-String
                                            }
                                        } -ErrorAction Stop
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }

                                    <#else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                            ''#>
                                }
                            }
                        }
                        '30wwk' {
                            # #30wwk
                            # Kill WOOFware sessions stuck on splashscreen; *could be improved
                            Invoke-Command -ComputerName $NetServices {
                                try {
                                    $WoofWareProcess = Get-Process -Name VCA.Sparky.Shell -IncludeUserName -ErrorAction Stop
                                    $WoofWareProcess | Select-Object -Property WS, UserName, ProcessName, StartTime, ID, Handles | Sort-Object -Property Handles |
                                        Format-Table -Property @{n='Computer';e={$env:COMPUTERNAME}}, Username, @{n='MemoryKB';e={($_.WS/1KB)}}, ProcessName, StartTime, Handles, ID -AutoSize | Out-String

                                    $WoofWareProcess | Where-Object { $_.Handles -le 899 -and (((Get-Date) - $_.StartTime) -ge (New-TimeSpan -Seconds 15)) } |
                                        Stop-Process -Confirm:$true -Verbose
                                }
                                catch {
                                    Write-Warning "[$env:COMPUTERNAME] $($PSItem.Exception.Message)"
                                }
                            }
                        }
                        '31' {
                            # #31
                            # Check UPS
                            if ((-not $Cluster) -and $NetServices) {
                                Get-UpsStatus -UPSs $(Convert-VcaAU -AU $ComputerName -Suffix '-ups') -ErrorAction SilentlyContinue |
                                Select-Object -ExcludeProperty PSComputerName, PSSourceJobInstanceId | Out-ListString
                            }
                            elseif ($Cluster) {
                                # Cluster
                                Get-UpsStatus -UPSs $(($ComputerName -replace '-vm\d+','-ups01'),($ComputerName -replace '-vm\d+','-ups02')) -ErrorAction SilentlyContinue |
                                Select-Object -ExcludeProperty PSComputerName, PSSourceJobInstanceId | Out-ListString
                            }
                            else {
                                Get-UpsStatus -UPSs $ComputerName -ErrorAction SilentlyContinue |
                                Select-Object -ExcludeProperty PSComputerName, PSSourceJobInstanceId | Out-ListString
                            }
                        }
                        '31l' {
                            # #31L
                            # launch UPS reports directory
                            Invoke-Item -Path '\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Reporting\2018-03 UPS Info & Set Configuration\Reports'
                        }
                        '31r' {
                            # #31r
                            # Search UPS reports #beta
                            $UPSReports = Get-ChildItem -Path "\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Reporting\2018-03 UPS Info & Set Configuration\Reports\HPUPSStatus-Report*.csv" -Recurse |
                                Select-String "$(Convert-VcaAU -AU $ComputerName -Suffix '-ups')" | Select-Object -Property Pattern, LineNumber, Line, Path
                            $UPSReports | Out-GridView -Title "#31r Select drive report log(s) to launch - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple |
                                Select-Object -ExpandProperty Path | Invoke-Item
                        }
                        '31s' {
                            # #31s
                            # Query UPS via SNMP
                            if ((-not $Cluster) -and $NetServices) {
                                Get-UpsSnmp -UPS $(Convert-VcaAU -AU $ComputerName -Suffix '-ups') -SnmpFallback -HPUpsSnmp $HPUpsSnmp | Out-ListString
                            }
                            elseif ($Cluster) {
                                # Cluster
                                Get-UpsSnmp -UPS $(($ComputerName -replace '-vm\d+','-ups01'),($ComputerName -replace '-vm\d+','-ups02')) -SnmpFallback -HPUpsSnmp $HPUpsSnmp | Out-ListString
                            }
                            else {
                                Get-UpsSnmp -UPS $ComputerName -SnmpFallback -HPUpsSnmp $HPUpsSnmp | Out-ListString
                            }
                        }
                        '31w' {
                            # #31w
                            # Launch UPS page
                            if ((-not $Cluster) -and $NetServices) {
                                Start-Process "https://$(Convert-VcaAU -AU $ComputerName -Suffix '-ups')"
                            }
                            elseif ($Cluster) {
                                # Cluster
                                ($ComputerName -replace '-vm\d+','-ups01'),($ComputerName -replace '-vm\d+','-ups02') | ForEach-Object {
                                    Start-Process "https://$PSItem"
                                }
                            }
                        }
                        '32' {
                            # #32
                            # Launch iLO remote console
                            if (-not $IloCredential) { $IloCredential = Get-Credential -Message "iLO Credentials:" }
                            if ($IloCredential) {
                                $IloCredParams = "-name $($IloCredential.UserName) -password $($IloCredential.GetNetworkCredential().Password)"

                                if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                    New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                }
                            }
                            if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$') {
                                Start-Process "$PSScriptRoot\Private\bin\HPE iLO Integrated Remote Console\HPLOCONS.exe" -ArgumentList "-addr $($ComputerName.Replace('.vcaantech.com',''))-ilo $IloCredParams"
                            }
                            else {
                                Start-Process "$PSScriptRoot\Private\bin\HPE iLO Integrated Remote Console\HPLOCONS.exe" -ArgumentList "-addr $ComputerName $IloCredParams"
                            }
                        }
                        '33' {
                            # #33
                            # Check server UID status
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                Clear-Variable -Name IloObj -ErrorAction Ignore
                                if (-not $Cluster) { $ServerIlo = "$ComputerName-ilo" }
                                else {
                                    $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                                }

                                if (-not $IloCredential) { $IloCredential = Get-Credential -Message "iLO Credentials:" }

                                if ($IloCredential) {
                                    try {
                                        $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $IloCredential -DisableCertificateAuthentication -ErrorAction Stop

                                        $IloConnection | ForEach-Object {
                                            $IloConnection_Item = $PSItem
                                            $IloObj = Get-HPEiLOUIDStatus -Connection $IloConnection_Item

                                            (Write-Output $IloConnection_Item | Select-Object -Property Hostname, IP, ServerModel, ServerGeneration, @{n = 'IndicatorLED'; e = { $IloObj.IndicatorLED } } |
                                                Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''

                                            # Draw horizontal line for clusters but exclude drawing line on last host.
                                            if ($IloConnection.Count -gt 1 -and $IloConnection.IndexOf($IloConnection_Item) -lt ($IloConnection.Count - 1)) { Write-Host "$('-'*70)`r" -ForegroundColor Cyan }
                                        }
                                    }
                                    catch {
                                        Write-Warning $Error[0].Exception.Message
                                        Clear-Variable -Name IloCredential
                                        Write-Host ''
                                    }
                                    if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                        New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                    }
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '33t' {
                            # #33t
                            # Toggle server UID status
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                Clear-Variable -Name IloObj -ErrorAction Ignore
                                if ((-not $Cluster) -and $NetServices) { $ServerIlo = "$ComputerName-ilo" }
                                elseif ($Cluster) {
                                    $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                                }

                                if (-not $IloCredential) { $IloCredential = Get-Credential -Message 'iLO Credentials:' }

                                if ($IloCredential) {
                                    try {
                                        $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $IloCredential -DisableCertificateAuthentication -ErrorAction Stop

                                        $IloConnection | ForEach-Object {
                                            $IloConnection_Item = $PSItem
                                            $IloObj = Get-HPEiLOUIDStatus -Connection $IloConnection_Item

                                            (Write-Output $IloConnection_Item | Select-Object -Property Hostname, IP, ServerModel, ServerGeneration, @{n = 'IndicatorLED'; e = { $IloObj.IndicatorLED } } |
                                                Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''

                                            if (-not $Cluster -or ($Cluster -and ($IloConnection_Item.Hostname -eq "$ComputerName-ilo"))) {
                                                Write-Host 'Toggling UID...' -ForegroundColor Cyan
                                                if ($IloObj.IndicatorLED -eq 'Off') {
                                                    Set-HPEiLOUIDStatus -IndicatorLED Lit -Connection $IloConnection_Item
                                                }
                                                else {
                                                    Set-HPEiLOUIDStatus -IndicatorLED Off -Connection $IloConnection_Item
                                                }
                                                Start-Sleep -Seconds 1
                                                $IloObj = Get-HPEiLOUIDStatus -Connection $IloConnection_Item
                                                (Write-Output $IloObj | Select-Object -Property IndicatorLED | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''
                                            }

                                            # Draw horizontal line for clusters but exclude drawing line on last host.
                                            if ($IloConnection.Count -gt 1 -and $IloConnection.IndexOf($IloConnection_Item) -lt ($IloConnection.Count - 1)) { Write-Host "$('-'*70)`r" -ForegroundColor Cyan }
                                        }
                                        if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                            New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                        }
                                    }
                                    catch {
                                        Write-Warning $Error[0].Exception.Message
                                        Clear-Variable -Name IloCredential
                                        Write-Host ''
                                    }
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '34' {
                            # #34
                            # View IML
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                Clear-Variable -Name IloObj -ErrorAction Ignore
                                if ((-not $Cluster) -and $NetServices) { $ServerIlo = "$ComputerName-ilo" }
                                elseif ($Cluster) {
                                    $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                                }
                                if (-not $IloCredential) { $IloCredential = Get-Credential -Message 'iLO Credentials:' }

                                if ($IloCredential) {
                                    try {
                                        $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $IloCredential -DisableCertificateAuthentication -ErrorAction Stop
                                        $IloObj = Get-HPEiLOIML -Connection $IloConnection

                                        $IloObj | ForEach-Object {
                                            $IloObj_Item = $_
                                            $_.IMLLog | Select-Object -Property @{n='Hostname';e={$IloObj_Item.Hostname}}, @{n='Updated';e={[datetime]$_.Updated}}, Count, Severity, Message, @{n='Created';e={[datetime]$_.Updated}}

                                        } | Sort-Object -Property @{Expression = 'Hostname'; Descending = $False }, @{Expression = 'Updated'; Descending = $True } |
                                            Out-GridView -Title "#34 [$($ServerIlo -join ', ')] - Integrated Management Log - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"

                                        if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                            New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                        }
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                        Clear-Variable -Name IloCredential
                                        Write-Host ''
                                    }
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '35' {
                            # #35
                            # Launch Service Now with IT Cases (ignoring HR Cases)
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/now/nav/ui/classic/params/target/task_list.do%3Fsysparm_query%3Dsys_class_nameANYTHING%255Elocation.u_display_nameSTARTSWITH$SNOWFilter-%255EnumberNOT%2520LIKEHRC%255EnumberNOT%2520LIKEHRT%255EnumberNOT%2520LIKECR%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '35e' {
                            # #35e
                            # Launch Service Now
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3Du_departmentSTARTSWITH$SNOWFilter%2520-%255EORlocationSTARTSWITH$SNOWFilter%2520-%255EORshort_descriptionLIKEAU$SNOWFilter%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '35i' {
                            # #35i (The modified version of the original before-v.200820 #35)
                            # Launch Service Now
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3Du_departmentSTARTSWITH$SNOWFilter%2520-%255EORlocationSTARTSWITH$SNOWFilter%2520-%255E%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '35o' {
                            # #35o
                            # Launch Service Now; filter for operations queue
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3Du_departmentSTARTSWITH$SNOWFilter%2520-%5Eassignment_group%253D0d56ea0c3742ca0012caa6d2b3990e17%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '35oe' {
                            # #35oe
                            # Launch Service Now; filter for operations queue
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3Du_departmentSTARTSWITH$SNOWFilter%2520-%255EORlocationSTARTSWITH$SNOWFilter%2520-%255EORshort_descriptionLIKEAU$SNOWFilter%255Eassignment_group%253D0d56ea0c3742ca0012caa6d2b3990e17%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '35t' {
                            # #35t
                            # Launch Service Now; filter for operations queue
                            Clear-Variable -Name SNOWFilter -ErrorAction Ignore
                            $SNOWFilter = $(Convert-VcaAu -AU $ComputerName -Strip)
                            $SNOWURL = "https://marsvh.service-now.com/nav_to.do?uri=%2Ftask_list.do%3Fsysparm_query%3Dsys_class_name!%3Dvtb_task%5Esys_class_name!%3Dsn_hr_core_case_performance%5Esys_class_name!%3Dhr_task%5Esys_class_name!%3Dsn_hr_core_case_talent_management%5Esys_class_name!%3Dsn_hr_core_task%5ElocationSTARTSWITH$SNOWFilter-%26sysparm_first_row%3D1%26sysparm_view%3D"

                            Start-Process $SNOWURL
                            Write-Host "$SNOWURL`r`n" -ForegroundColor Cyan
                        }
                        '36' {
                            # #36
                            # Reset iLO
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                if ((-not $Cluster) -and $NetServices) { $ServerIlo = "$ComputerName-ilo" }
                                elseif ($Cluster) {
                                    $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                                }
                                if (-not $IloCredential) { $IloCredential = Get-Credential -Message 'iLO Credentials:' }

                                if ($IloCredential) {
                                    try {
                                        $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $IloCredential -DisableCertificateAuthentication -ErrorAction Stop
                                        Reset-HPEiLO -Connection $IloConnection -Device iLO | Out-String

                                        if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                            New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                        }
                                    }
                                    catch {
                                        Write-Warning $Error[0].Exception.Message
                                        #Clear-Variable -Name IloCredential
                                        Write-Host ''
                                    }
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '36a' {
                            # #36a
                            # Send Auxilary cycle signal
                            if ((-not $Cluster) -and $NetServices) {
                                $ServerIlo = "$ComputerName-ilo"
                            }
                            else {
                                $ServerIlo = $Cluster | ForEach-Object { "$($_ -replace '.vcaantech.com','')-ilo" }
                            }

                            if (-not $IloCredential) { $IloCredential = Get-Credential -Message "iLO Credentials:" }
                            if ($IloCredential) {
                                $Headers = @{ Authorization = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($IloCredential.UserName)`:$($IloCredential.GetNetworkCredential().Password)")) }

                                try {
                                    $IloRestConnection = @()
                                    $ServerIlo | ForEach-Object {
                                        $IloRestConnection += Invoke-RestMethod -Uri "https://$_/redfish/v1/" -Method Get -Headers $Headers -ContentType "application/json" -SessionVariable "IloRestSession-$_" -ErrorAction Stop
                                    }
                                    $IloRestConnection | ForEach-Object {
                                        $IloRestConnection_Item = $_
                                        $IloHostName = ($IloRestConnection_Item.Oem.Hpe.Manager, $IloRestConnection_Item.Oem.Hp.Manager | Where-Object { $_.HostName }).HostName
                                        $IloGeneration = ($IloRestConnection_Item.Oem.Hpe.Manager, $IloRestConnection_Item.Oem.Hp.Manager | Where-Object { $_.ManagerType }).ManagerType

                                        $AuxPowerCylceBody = @{
                                            Action    = 'SystemReset'
                                            ResetType = 'AuxCycle'
                                        } | ConvertTo-Json

                                        if ((Read-Choice -Title "#36a Send auxilary cycle action - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message "$IloHostName ($IloGeneration) - Aux cycle?") -eq 1) {
                                            try {
                                                $IloAuxCycleUri = switch ($IloGeneration) {
                                                    'iLO 5' { 'redfish/v1/Systems/1/Actions/Oem/Hpe/HpeComputerSystemExt.SystemReset/'; break }
                                                    'iLO 4' { 'redfish/v1/Systems/1/Actions/Oem/Hp/ComputerSystemExt.SystemReset/'; break }
                                                    default { 'redfish/v1/Systems/1/Actions/Oem/Hpe/HpeComputerSystemExt.SystemReset/' }
                                                }
                                                $IloRestParams = @{
                                                    Uri         = "https://$IloHostName/$IloAuxCycleUri"
                                                    Method      = 'Post'
                                                    Body        = $AuxPowerCylceBody
                                                    ContentType = "application/json"
                                                    WebSession  = (Get-Variable -Name "IloRestSession-$IloHostName").Value
                                                }
                                                $IloResult = Invoke-RestMethod @IloRestParams -ErrorAction Stop

                                                $IloRestOutput = [pscustomobject]@{
                                                    Hostname      = $IloHostName
                                                    iLOGeneration = $IloGeneration
                                                    ActionMessage = $IloResult.error.'@Message.ExtendedInfo'.MessageId
                                                }
                                                $IloRestOutput | Out-TableString

                                                Write-Host "The auxiliary power-cycle will take place the next time that the server is rebooted or powered off.`r`n" -ForegroundColor Cyan

                                                if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                                    New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                                }
                                            }
                                            catch {
                                                Write-Warning $_.Exception.Message
                                            }
                                        }
                                        else {
                                            Write-Host ''
                                        }
                                    }
                                }
                                catch {
                                    Write-Warning $_.Exception.Message
                                    #Clear-Variable -Name IloCredential
                                    Write-Host ''
                                }
                                Clear-Variable -Name IloRestSession-* -ErrorAction Ignore
                            }
                        }
                        '36a3' {
                            # #36a3
                            # Send Auxilary cycle signal
                            if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
                                if ((-not $Cluster) -and $NetServices) {
                                    $ServerIlo = "$ComputerName-ilo"
                                }
                                else {
                                    $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
                                }

                                if (-not $IloCredential) { $IloCredential = Get-Credential -Message "iLO Credentials:" }

                                if ($IloCredential) {
                                    try {
                                        $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $IloCredential -DisableCertificateAuthentication -ErrorAction Stop
                                        $IloConnection | ForEach-Object {
                                            $IloConnection_Item = $_

                                            $Headers = @{ Authorization = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($IloCredential.UserName)`:$($IloCredential.GetNetworkCredential().Password)")) }
                                            $AuxPowerCylceBody = @{
                                                Action    = 'SystemReset'
                                                ResetType = 'AuxCycle'
                                            } | ConvertTo-Json

                                            if ((Read-Choice -Title "#36a3 Send auxilary cycle action - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message "$($IloConnection_Item.Hostname) - Aux cycle?") -eq 1) {
                                                try {
                                                    $IloAuxCycleUri = switch ($IloConnection_Item.TargetInfo.iLOGeneration) {
                                                        'iLO5' { 'redfish/v1/Systems/1/Actions/Oem/Hpe/HpeComputerSystemExt.SystemReset/'; break }
                                                        'iLO4' { 'redfish/v1/Systems/1/Actions/Oem/Hp/ComputerSystemExt.SystemReset/'; break }
                                                        default { 'redfish/v1/Systems/1/Actions/Oem/Hpe/HpeComputerSystemExt.SystemReset/' }
                                                    }
                                                    $IloResult = Invoke-RestMethod -Uri "https://$($IloConnection_Item.Hostname)/$IloAuxCycleUri" -Method Post -Headers $Headers -Body $AuxPowerCylceBody -ContentType "application/json" -ErrorAction Stop

                                                    $IloRestOutput = [pscustomobject]@{
                                                        Hostname      = $IloConnection_Item.Hostname
                                                        ActionMessage = $IloResult.error.'@Message.ExtendedInfo'.MessageId
                                                    }
                                                    $IloRestOutput | Out-TableString

                                                    Write-Host "The auxiliary power-cycle will take place the next time that the server is rebooted or powered off.`r`n" -ForegroundColor Cyan

                                                    if ($IloCredential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                                                        New-StoredCredential -Credentials $IloCredential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                                                    }
                                                }
                                                catch {
                                                    Write-Warning $_.Exception.Message
                                                }
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                        #Clear-Variable -Name IloCredential
                                        Write-Host ''
                                    }
                                }
                            }
                            else {
                                Write-Warning "HPEiLOCmdlets module not found."
                                Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
                            }
                        }
                        '37' {
                            # #37
                            # WW API Health Report
                            if (-not $SNOWAPICredential) { $SNOWAPICredential = Get-StoredCredential -Target vcasnowapi }
                            Get-WWAPIHealth -Credential $SNOWAPICredential -HospitalMaster $HospitalMaster
                        }
                        '37c' {
                            # #37c
                            Update-HospitalMaster -EmailCredential $EmailCredential
                            $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc

                            Invoke-WWAPIVerification -HospitalMaster $HospitalMaster
                        }
                        '37i'{
                            # #37i
                            #Get API check for single site
                            $APISINGLEURL = "https://api.vcaantech.com/api/v1/hospitals/$(Convert-VcaAu -AU $ComputerName -Strip)/health/check"
                            Start-Process $APISINGLEURL
                            Write-Output $APISINGLEURL
                        }
                        '37rdp' {
                            # #37rdp
                            # scan rdp network port
                            Clear-Variable -Name RDPServicesResults, RDPServicesCase -ErrorAction Ignore
                            $HospitalNS = (Get-ADComputer -Filter '(Name -like "h*-ns*" -and Name -notlike "*-old" -and Name -notlike "h8*-ns*") -and OperatingSystem -like "*Server*" -and Enabled -eq $true' |
                                Select-Object -ExpandProperty Name) -match '^h\d+-ns\d{0,2}$' | Sort-Object

                            $HospitalNS | Start-RSJob -Name RdpPortJobs -FunctionsToImport Test-ConnectionAsync, Convert-VcaAU -VariablesToImport HospitalMaster -Throttle 64 -ScriptBlock {
                                try { $RdpPortResponse = [System.Net.Sockets.TcpClient]::new().ConnectAsync("$_", 3389).Wait(10000) }
                                catch { $ErrorMessage = $_.Exception.Message }

                                if ($RdpPortResponse -in $false, $null) {
                                    $ComputerName_Item = $_
                                    $InHospitalMaster = $(
                                            if ($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq "$(Convert-VcaAU -AU $ComputerName_Item -Strip)" } )) { 'Yes' }
                                            else { '--' }
                                    )
                                    $PingResponse = $(
                                        $ErrorActionPreference = 'Ignore'
                                        if (($RdpPing = Test-ConnectionAsync -ComputerName $_ -Full).ResponseTime) {
                                            "$($RdpPing.ResponseTime)"
                                        }
                                        else { '--' }
                                        $ErrorActionPreference = 'Continue'
                                    )
                                }
                                [pscustomobject]@{
                                    ComputerName            = $_
                                    'RdpPortResponse(3389)' = $RdpPortResponse
                                    InHospitalMaster        = $InHospitalMaster
                                    'PingResponse (ms)'     = $PingResponse
                                    Error                   = $ErrorMessage
                                }
                            } | Out-Null
                            $RDPServicesResults = Get-RSJob -Name RdpPortJobs | Wait-RSJob -ShowProgress -Timeout 600 | Receive-RSJob | Sort-Object -Property 'RdpPortResponse(3389)'
                            Get-RSJob -Name RdpPortJobs | Remove-RSJob -Force
                            $RDPServicesResults | Out-GridView -Title "#37rdp Global Remote Desktop Services Response (Port: 3389) - Select site(s) to generate ServiceNow Case - $((Get-Date).ToString("yyyy-MM-dd HH:mm")) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable RDPServicesCase | Out-Null

                            # Generate ServiceNow Incident
                            if ($RDPServicesCase) {
                                # Get user's email address using .NET
                                $ADQuery = New-Object System.DirectoryServices.DirectorySearcher
                                $ADQuery.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
                                $ADQuery.Filter = "(&(ObjectClass=User)(samAccountName=$env:USERNAME))"
                                $ADQueryResult = $ADQuery.FindOne()

                                $MsgBoxInput = [System.Windows.MessageBox]::Show('Are you sure?', 'Generate ServiceNow Incident', 'YesNo', 'Question')
                                switch ($MsgBoxInput) {
                                    'Yes' {
                                        if (-not $SNOWAPICredential) { $SNOWAPICredential = Get-StoredCredential -Target vcasnowapi }
                                        $RDPServicesCase | ForEach-Object {
                                            $RdpSiteAU = Convert-VcaAU -AU $PSItem.ComputerName -Strip -ErrorAction SilentlyContinue
                                            $RdpSiteTimeZone = $(($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq $RdpSiteAU } )).'Time Zone')
                                            $RdpShortDescription = "AU$RdpSiteAU - $($PSItem.ComputerName) - Remote Desktop Service Unreachable - TZ: $RdpSiteTimeZone"

                                            Clear-Variable -Name NewSnowCaseResult -ErrorAction Ignore
                                            $NewServiceNowParams = @{
                                                ImpactedUser     = $(($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq $RdpSiteAU } )).'Hospital Manager Email') #String email address of SNOW user:Default Guest
                                                ReportedBy       = $($ADQueryResult.Properties.mail) #String email address of SNOW user:Default Guest

                                                Category         = "Network"               #String accepted values:"hardware";"inquiry";"network";"software";"database";"security":Default "inquiry"
                                                CIName           = $($PSItem.ComputerName) #String:Name of CI:Default "Hardware - Other"
                                                Impact           = "1"                     #Numeric:Range 1 to 3:Default 3
                                                Urgency          = "2"                     #Numeric:Range 1 to 3:Default 3

                                                ContactType      = "Monitoring"     #String: accepted values: "messenger";"email";"phone";"self-service";"monitoring";"voice mail";"walk-in/direct-contact":DEFAULT "monitoring"
                                                AssignedGroup    = "VCA Operations" #String:Name of SNOW Group:Default "Support Alerts"
                                                AssignedTo       = ""               #Name of of a member of SNOW group:Default NULL if AssignedTo user is not a member of the AssignedGroup

                                                ShortDescription = "$RdpShortDescription" #String: Default value "Short Description is MISSING"
                                                Description      = "Remote Desktop Service Report: $($PSItem | Format-List | Out-String)" #String: Default value "Description is MISSING"
                                            }
                                            $NewServiceNowParams.Credential = $SNOWAPICredential
                                            $NewSnowCaseResult = New-ServiceNowIncident @NewServiceNowParams

                                            if ($NewSnowCaseResult) {
                                                Write-Host "`r`nCase Generated: $($NewSnowCaseResult.incident) - $RdpShortDescription`r`n" -ForegroundColor Yellow
                                                $PSItem | Out-ListString
                                                Write-Host "`r`n$($NewSnowCaseResult.url)" -ForegroundColor Cyan
                                                Start-Process $NewSnowCaseResult.url
                                            }
                                            Write-Host ''
                                        }
                                        break
                                    }
                                    'No' {
                                    }
                                } #MsgBox Switch Case
                            } #generate snow case
                        }
                        '38' {
                            # #38
                            # Query installed applications
                            Clear-Variable -Name SiteServers, InstalledApps -ErrorAction Ignore
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#38 Select server to query installed applications - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                            # Proceed if server was selected
                            if ($SiteServers) {
                                $InstalledApps = Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock {
                                    Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' |
                                        Select-Object -Property DisplayName, InstallDate, InstallLocation, DisplayVersion, Publisher, PSParentPath | Sort-Object -Property InstallDate -Descending
                                } -ErrorAction SilentlyContinue

                                $InstalledApps | Select-Object -Property PSComputerName, DisplayName, InstallDate, InstallLocation, DisplayVersion, Publisher, PSParentPath |
                                    Out-GridView -Title "#38 Installed applications - Send Results to Console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -PassThru | Out-String
                            }
                        }
                        '38l' {
                            # #38l
                            # MsiInstallation application log
                            Clear-Variable -Name SiteServers, InstalledApps -ErrorAction Ignore
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#38L Select server to query MSIinstallation log - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                            # Proceed if server was selected
                            if ($SiteServers) {
                                $scriptblock = {Get-EventLog -LogName application -Source msiinstaller   |
                                    Select-Object -Property username,Index,EntryType,TimeWritten,Source,EventID,Message
                                }
                                Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock $scriptblock |
                                    Select-Object -Property username,Index,EntryType,TimeWritten,Source,EventID,Message,PSCOmputerName |
                                    Sort-Object -Property TimeWritten -Descending | 
                                    Out-GridView  -Title "#38L Application MsiInstaller Log - $((Get-Date).ToString("yyyy-MM-dd HH:mm")) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"-PassThru
                                }
                        }
                        '40' {
                            # #40
                            # PortQry.exe
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#40 Select Windows machine(s) to query network ports - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                                [System.Management.Automation.Runspaces.PSSession[]]$PSSession = @()

                                $SiteServers | foreach-object {
                                    $SiteServers_Item = $PSItem
                                    try {
                                        $PSSession = $PSSession + (New-PSSession -ComputerName $PSItem.Name -ErrorAction Stop)
                                    }
                                    catch {
                                        Write-Warning "[$($SiteServers_Item.Name)] $($_.Exception.Message)"
                                        Write-Host ''
                                    }
                                }
                                $PSSession | ForEach-Object {
                                    Copy-ToPSSession -Path "$PSScriptRoot\Private\bin\PortQry.exe" -Destination "C:\Windows\System32" -Session $PSItem
                                    Write-Host "`r`nRemote Machine: $($PSItem.ComputerName)`r`n" -ForegroundColor Cyan
                                    $ScriptBlock = {
                                        Write-Host "PortQry.exe -n www.google.com -p both -o 80,443"
                                        Write-Host "PortQry.exe -n phrdslp01 -p both -o 5985,49298"
                                        Write-Host "PortQry.exe -n api.vcaantech.com -p both -o 443"

                                        PortQry.exe -n www.google.com -p both -o 80`,443
                                        PortQry.exe -n phrdslp01 -p both -o 5985`,49298
                                        PortQry.exe -n api.vcaantech.com -p both -o 443
                                    }
                                    Invoke-Command -Session $PSItem -ScriptBlock $ScriptBlock

                                    if ($PSSession.Count -gt 1 -and $PSSession.IndexOf($PSItem) -lt ($PSSession.Count - 1)) { Write-Host "$('-'*55)" }

                                }
                                if ($PSSession) { Remove-PSSession -Session $PSSession }
                                Write-Host ''
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '40s' {
                            # #40s
                            # PortQry.exe
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#40s Select Windows machine(s) to query network ports - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                                [System.Management.Automation.Runspaces.PSSession[]]$PSSession = @()

                                $SiteServers | foreach-object {
                                    $SiteServers_Item = $PSItem
                                    try {
                                        $PSSession = $PSSession + (New-PSSession -ComputerName $PSItem.Name -ErrorAction Stop)
                                    }
                                    catch {
                                        Write-Warning "[$($SiteServers_Item.Name)] $($_.Exception.Message)"
                                        Write-Host ''
                                    }
                                }
                                $PSSession | ForEach-Object {
                                    Copy-ToPSSession -Path "$PSScriptRoot\Private\bin\PortQry.exe" -Destination "C:\Windows\System32" -Session $PSItem
                                    Write-Host "`r`nRemote Machine: $($PSItem.ComputerName)`r`n" -ForegroundColor Cyan
                                    $ScriptBlock = {

                                        Write-Host "PortQry.exe -n PHVCAPMP01 -p TCP -o 52311"
                                        Write-Host "PortQry.exe -n 10.125.110.151 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.125.110.152 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.152 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.153 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.154 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.155 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.156 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.157 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n 10.230.107.158 -p TCP -o 9997"
                                        Write-Host "PortQry.exe -n PHISSPDepP01 -p TCP -o 8089"
                                        Write-Host "PortQry.exe -n LAVCCATRENDD01 -p both -o 29401,46365"
                                        Write-Host "PortQry.exe -n LAVCCATRENDD01 -p TCP -o 4343,8080"
                                        Write-Host "PortQry.exe -n PHISNGAVP01 -p both -o 29401,46365"
                                        Write-Host "PortQry.exe -n PHISNGAVP01 -p TCP -o 4343,8080"
                                        Write-Host "PortQry.exe -n PHISPKIOInfraP1 -p TCP -o 80,443,135"
                                        Write-Host "PortQry.exe -n PHISPKIOInfraP2 -p TCP -o 80,443,135"
                                        Write-Host "PortQry.exe -n PHISPKIWInfraP1 -p TCP -o 80,443,135"
                                        Write-Host "PortQry.exe -n PHISPKIWInfraP2 -p TCP -o 80,443,135"
                                        Write-Host "PortQry.exe -n PHISPKICInfraP1 -p TCP -o 135,6212"

                                        PortQry.exe -n PHVCAPMP01 -p TCP -o 52311
                                        PortQry.exe -n 10.125.110.151 -p TCP -o 9997
                                        PortQry.exe -n 10.125.110.152 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.152 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.153 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.154 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.155 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.156 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.157 -p TCP -o 9997
                                        PortQry.exe -n 10.230.107.158 -p TCP -o 9997
                                        PortQry.exe -n PHISSPDepP01 -p TCP -o 8089
                                        PortQry.exe -n LAVCCATRENDD01 -p both -o 29401`,46365
                                        PortQry.exe -n LAVCCATRENDD01 -p TCP -o 4343`,8080
                                        PortQry.exe -n PHISNGAVP01 -p both -o 29401`,46365
                                        PortQry.exe -n PHISNGAVP01 -p TCP -o 4343`,8080
                                        PortQry.exe -n PHISPKIOInfraP1 -p TCP -o 80`,443`,135
                                        PortQry.exe -n PHISPKIOInfraP2 -p TCP -o 80`,443`,135
                                        PortQry.exe -n PHISPKIWInfraP1 -p TCP -o 80`,443`,135
                                        PortQry.exe -n PHISPKIWInfraP2 -p TCP -o 80`,443`,135
                                        PortQry.exe -n PHISPKICInfraP1 -p TCP -o 135`,6212
                                    }
                                    Invoke-Command -Session $PSItem -ScriptBlock $ScriptBlock

                                    if ($PSSession.Count -gt 1 -and $PSSession.IndexOf($PSItem) -lt ($PSSession.Count - 1)) { Write-Host "$('-'*55)" }

                                }
                                if ($PSSession) { Remove-PSSession -Session $PSSession }
                                Write-Host ''
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '40st' {
                            # #40st
                            # PortQry.exe
                            if (Get-Module -Name ActiveDirectory) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#40st Select Windows machine to perform internet speed test - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single
                                [System.Management.Automation.Runspaces.PSSession[]]$PSSession = @()

                                $SiteServers | foreach-object {
                                    $SiteServers_Item = $PSItem
                                    try {
                                        $PSSession = $PSSession + (New-PSSession -ComputerName $PSItem.Name -ErrorAction Stop)
                                    }
                                    catch {
                                        Write-Warning "[$($SiteServers_Item.Name)] $($_.Exception.Message)"
                                        Write-Host ''
                                    }
                                }
                                $PSSession | ForEach-Object {
                                    Copy-ToPSSession -Path "$PSScriptRoot\Private\bin\speedtest.exe" -Destination "C:\temp" -Session $PSItem
                                    Write-Host "`r`nRemote Machine: $($PSItem.ComputerName)`r`n" -ForegroundColor Cyan

                                    try {
                                        $ScriptBlock = {
                                            & 'C:\temp\speedtest.exe' --accept-license
                                            Write-Host ''
                                        }
                                        Invoke-Command -Session $PSItem -ScriptBlock $ScriptBlock -ErrorAction Stop
                                    }
                                    catch {
                                        if ($_.Exception.Message -notlike "WARNING: =*") {
                                            #Write-Warning $_.Exception.Message
                                        }
                                        Write-Host ''
                                    }
                                }
                                if ($PSSession) { Remove-PSSession -Session $PSSession }
                                Write-Host ''
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '41' {
                            # #41
                            # Find temp folders
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                $NetServices | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        icacls \\$PSItem\c$\users\temp*
                                    } #PingStatus
                                } #ForEach
                            }
                        }
                        '41d' {
                            # #41d
                            #VHD Mounts
                            whatdisk -ComputerName $NetServices |
                                Out-GridView -Title "#41d AU$(Convert-VcaAu -AU $ComputerName -Strip) - VHD Mounts - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple | Out-String
                        }
                        '41l' {
                            # #41L
                            # find/delete orphan user junction links
                            if (@($NetServices).count -le 1) {
                                Write-Host ''
                                Write-Warning "Orphan Junction Link purge is only for Multi-NS sites with VHDX User Profiles."
                                Write-Host ''
                                break
                            }
                            Clear-Variable -Name NetServicesOnline, JunctionLinksDeletion, JunctionResults -ErrorAction Ignore
                            $NetServicesOnline = $Netservices | ForEach-Object { if ($_ | Get-PingStatus) { $_ }}

                            $ScriptBlock = {
                                foreach ($JunctionLink_Item in (Get-ChildItem -Path 'C:\Users' | Where-Object LinkType -EQ Junction)) {
                                    Get-ChildItem -Path $JunctionLink_Item.FullName -ErrorAction SilentlyContinue | Out-Null
                                    if (-not $?) {
                                        if ($Error[0].Exception.GetType().FullName -eq 'System.IO.DirectoryNotFoundException') {
                                            [pscustomobject]@{
                                                FullName       = $JunctionLink_Item.FullName
                                                LinkType       = $JunctionLink_Item.LinkType
                                                CreationTime   = $JunctionLink_Item.CreationTime
                                                LastAccessTime = $JunctionLink_Item.LastAccessTime
                                            }
                                        }
                                    }
                                } #foreach
                            } #scriptblock
                            $JunctionResults = Invoke-Command -ComputerName $NetServicesOnline -ScriptBlock $ScriptBlock | Select-Object -Property PSComputerName, FullName, LinkType, CreationTime, LastAccessTime
                            $JunctionResults | Out-GridView -Title "#41L Orphan Junction Links - Select and verify for deletion - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable JunctionLinksDeletion | Out-Null

                            if ($JunctionLinksDeletion) {
                                Write-Host "`r`nDeleting Orphan User Profile Junction Links:" -ForegroundColor Cyan

                                $JunctionLinksDeletion | Group-Object -Property PSComputerName | ForEach-Object {
                                    Clear-Variable -Name JunctionLinksDeletion_Item -ErrorAction Ignore
                                    $JunctionLinksDeletion_Item = $_

                                    Write-Host "`r`n$($_.Name):" -ForegroundColor Cyan
                                    Invoke-Command -ComputerName $_.Name -ScriptBlock {
                                        Remove-Item -Path $using:JunctionLinksDeletion_Item.Group.FullName -Force -Verbose -Confirm:$false -ErrorAction SilentlyContinue
                                    }
                                } #foreach
                                Write-Host ''
                            }
                            else {
                                Write-Host "`r`n[$SiteAUNumber] No orphan User Profile Junction Links found`r`n" -ForegroundColor Cyan
                            }
                        }
                        '41lg' {
                            # #41Lg
                            # find/delete orphan user junction links
                            Clear-Variable -Name NetServicesOnline, JunctionLinksDeletion, JunctionResults -ErrorAction Ignore
                            $MultiNSSites = (Get-ADComputer -Filter '((Name -like "h*-ns*" -and Name -notlike "h*-ns" -and Name -notlike "h8*-ns*") -and Name -notlike "*-old") -and OperatingSystem -like "*Server*" -and Enabled -eq $true').Name -match '^h\d+-ns\d+$' | Sort-Object
                            $NetServicesOnline = (Test-ConnectionAsync -ComputerName $MultiNSSites -ErrorAction SilentlyContinue | Where-Object Result -eq 'Success').ComputerName

                            $ScriptBlock = {
                                foreach ($JunctionLink_Item in (Get-ChildItem -Path 'C:\Users' | Where-Object LinkType -EQ Junction)) {
                                    Get-ChildItem -Path $JunctionLink_Item.FullName -ErrorAction SilentlyContinue | Out-Null
                                    if (-not $?) {
                                        if ($Error[0].Exception.GetType().FullName -eq 'System.IO.DirectoryNotFoundException') {
                                            [pscustomobject]@{
                                                FullName       = $JunctionLink_Item.FullName
                                                LinkType       = $JunctionLink_Item.LinkType
                                                CreationTime   = $JunctionLink_Item.CreationTime
                                                LastAccessTime = $JunctionLink_Item.LastAccessTime
                                            }
                                        }
                                    }
                                } #foreach
                            } #scriptblock
                            $NetServicesOnline | Start-RSJob -Name JunctionLinkJobs -Throttle 64 -ScriptBlock {
                                    Invoke-Command -ComputerName $_ -ScriptBlock $using:ScriptBlock -Credential $using:ADCredential -ErrorAction SilentlyContinue |
                                        Select-Object -Property PSComputerName, FullName, LinkType, CreationTime, LastAccessTime
                            } | Out-Null
                            $JunctionResults = Get-RSJob -Name JunctionLinkJobs | Wait-RSJob -ShowProgress -Timeout 300 | Receive-RSJob
                            $JunctionResults | Out-GridView -Title "#41LG Orphan Junction Links - Select and verify for deletion - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable JunctionLinksDeletion | Out-Null

                            if ($JunctionLinksDeletion) {
                                Write-Host "`r`nDeleting Orphan User Profile Junction Links:" -ForegroundColor Cyan

                                $JunctionLinksDeletion | Group-Object -Property PSComputerName | ForEach-Object {
                                    Clear-Variable -Name JunctionLinksDeletion_Item -ErrorAction Ignore
                                    $JunctionLinksDeletion_Item = $_

                                    Write-Host "`r`n$($_.Name):" -ForegroundColor Cyan
                                    Invoke-Command -ComputerName $_.Name -ScriptBlock {
                                        Remove-Item -Path $using:JunctionLinksDeletion_Item.Group.FullName -Force -Verbose -Confirm:$false -ErrorAction SilentlyContinue
                                    }
                                } #foreach
                                Write-Host ''
                            }
                            else {
                                Write-Host "`r`nGlobal: No orphan User Profile Junction Links found`r`n" -ForegroundColor Cyan
                            }
                            Get-RSJob -Name JunctionLinkJobs | Remove-RSJob -Force
                        }
                        '41m' {
                            # #41m
                            # Find User folders on NS C:\users\
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                $NetServices | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        Write-Host $PSItem
                                        Invoke-Command -ComputerName $PSItem -ScriptBlock {
                                            Get-ChildItem -Path "$env:SystemDrive\users" |
                                                Select-Object -Property LinkType, Name, LastWriteTime |
                                                Sort-Object -Property LinkType, Name | Out-String
                                        }
                                    } #PingStatus
                                } #ForEach
                            }
                        }
                        '41n' {
                            # #41n
                            # Count Notifications in Registry
                            if (-not $NetServices) { break }
                            if ($NetServices -and $Cluster) {
                                Clear-Variable -Name ParallelTask, NotificationsRemove -ErrorAction Ignore
                                $ParallelTask = Invoke-Command -ComputerName $NetServices -ScriptBlock {
                                    Get-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Notifications" |
                                    Select-Object -Property PSComputerName, @{n = 'NotificationsCount'; e = { $_.Property.Count } }
                                } -AsJob

                                $ParallelTask | Wait-Job | Receive-Job -Keep -ErrorAction SilentlyContinue | Select-Object -Property @{n='Name';e={$_.PSComputerName}}, NotificationsCount |
                                    Out-GridView -OutputMode Multiple -Title "#41n Notifications Registry Entries - Select computer(s) to clear entries - HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Notifications - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutVariable NotificationsRemove |
                                    Format-Table -AutoSize | Out-String
                                if ($ParallelTask) { $ParallelTask | Remove-Job -Force }
                                if ($NotificationsRemove) {
                                    $NotificationsRemove | ForEach-Object {
                                        if ((Read-Choice -Title "#41n [$($PSItem.Name)] Notifications - HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Notifications - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message 'Clear notifications registry?') -eq 1) {
                                            Write-Host "[$($PSItem.Name)] Clearing notifications registry entries..." -ForegroundColor Cyan
                                            Invoke-Command -ComputerName $PSItem.Name -ScriptBlock {
                                                Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Notifications" -Recurse -Verbose
                                                New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Notifications" -Verbose
                                            }
                                            Write-Host "* [$($PSItem.Name)] Please reboot server to complete notifications purge." -ForegroundColor Magenta
                                        } #Read-Choice
                                    } #foreach-object
                                }
                            }
                            if (-not $Cluster) {
                                Write-Warning "Please use only on cluster NS sites and only if necessary."
                            }
                        }
                        '42' {
                            # #42
                            # Stop appreadiness
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Invoke-Command -ComputerName $NetServices -ScriptBlock {
                                    ($AppReadiStatus = Get-Service -Name AppReadiness)
                                    if ($AppReadiStatus.Status -ne 'Stopped') {
                                        Stop-Service -Name AppReadiness -Verbose
                                        Get-Service -Name AppReadiness
                                    }
                                }
                            }
                        }
                        '43' {
                           # #43
                            # Check for .bak entries in registry
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                if ((Read-Choice -Title "#43 [$Computername] Query registry ProfileList for .bak entries? - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -DefaultChoice 1) -eq 1) {
                                    $NetServices | ForEach-Object {
                                        if ($PSItem | Get-PingStatus) {
                                            Write-Host "Trying host $PSItem. Blank equals no .bak found"
                                            Invoke-command -ComputerName $PSItem -ScriptBlock { Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' } |
                                                Where-Object Name -Like "*.bak*" | Select-Object -Property Name, PSComputerName | Format-List | Out-String
                                        } #PingStatus
                                    } #ForEach
                                }
                            }
                        }

                        '43d' {
                            # #43d
                            # Check / Delete *.backup- folders in C:\users\
                            if (-not $NetServices) { break }
                            Clear-Variable -Name Removebackup -ErrorAction Ignore
                            if ($NetServices) {
                                $Removebackup = Invoke-Command -ComputerName $NetServices -ScriptBlock { Get-Item -Path "C:\Users\*.backup-*" } -ErrorAction SilentlyContinue |
                                    Select-Object -Property Mode, Name, LastWriteTime, PSComputerName |
                                    Out-GridView -PassThru -Title "#43d AU$(Convert-VcaAu -AU $ComputerName -Strip) Delete ALL User -Backup folders - NOTE: Verify Output -- Action CANNOT be undone - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                            }
                            if ($Removebackup) {
                                Invoke-Command -ComputerName $NetServices { Get-Item -Path "C:\Users\*.backup-*" | Remove-Item -Force -Recurse -Verbose -Confirm:$false }
                            }
                        }
                        '43f' {
                            # #43f
                            # Delete .bak entries in registry
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                $NetServices | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        Write-Warning "Server reboot is strongly recommended after running to prevent further issues."
                                        if ((Read-Choice -Title "#43f [$PSItem] Registry .bak profile entries - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message 'Delete .bak profile entries?') -eq 1) {
                                            Remove-BakRegistry -servers $PSItem
                                        }
                                    } #PingStatus
                                } #ForEach
                            }
                        }
                        '43p' {
                            # #43p
                            # Delete old user profiles
                            if (-not $script:ADCredential) { $script:ADCredential = Get-StoredCredential -Target vcadomaincreds }

                            $SiteProfilesSelection = Select-VcaSite -AU $ComputerName -Title "#43p Select server to query old profiles - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single
                            if (-not ($SiteProfilesSelection.Name -like "*-fs*")) {
                                if (-not $SiteProfilesSelection.Name -and $SingleHost) { $SiteProfilesSelection = @{Name = $ComputerName} }
                                if ($SiteProfilesSelection.Name) {
                                    try {
                                        $OldUserProfiles = Get-OldUserProfiles -ComputerName $SiteProfilesSelection.Name -Credential $ADCredential -ErrorAction Stop
                                        if ($OldUserProfiles) {
                                            $OldUserProfiles | Out-GridView -Title "#43p User Profile Purge - Count: $(@($OldUserProfiles).Count) - NOTE: Verify Output -- Action CANNOT be undone - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable UserProfileDeletion | Out-Null
                                            if ($UserProfileDeletion) {
                                                Write-Host "`r`nProfiles to delete:" -ForegroundColor Cyan
                                                $OldUserProfiles | Out-TableString

                                                Invoke-Command -ComputerName $SiteProfilesSelection.Name {
                                                    Get-CimInstance -ClassName Win32_UserProfile -Filter "Loaded=False And Special=False" | Where-Object SID -in $using:OldUserProfiles.SID | Remove-CimInstance -Verbose
                                                } -Credential $ADCredential
                                            }
                                        }
                                        else {
                                            Write-Host "[$($SiteProfilesSelection.Name)] No user profiles found for deletion" -ForegroundColor Cyan
                                        }
                                        Clear-Variable -Name OldUserProfiles -ErrorAction Ignore
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                }
                            } # if (-not ($SiteProfilesSelection.Name -like "*-fs*")) {
                            elseif ($SiteProfilesSelection.Name -like "*-fs*") {
                                Clear-Variable -Name FileServer -ErrorAction Ignore
                                $FileServer = (Get-ADComputer -Filter "Name -like '$(Convert-VcaAU -AU $ComputerName -Suffix -fs)*'").Name
                                $FileServer | ForEach-Object {
                                    if ($PSItem | Get-PingStatus) {
                                        try {
                                            $OldVhds = Get-OldVhds -ComputerName $PSItem -ErrorAction Stop
                                            Write-Host ''
                                            if ($OldVhds) {
                                                $VhdTotalSize = ($OldVhds | Measure-Object -Property Length -Sum).Sum
                                                $OldVhds | Out-GridView -Title "#43p Total VHD space to be deleted: $([math]::Round(($VhdTotalSize / 1GB), 2)) GB - NOTE: Verify Output -- Action CANNOT be undone - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable VhdDeletion | Out-Null
                                                if ($VhdDeletion) {
                                                    Write-Host "`r`nProfiles to delete:" -ForegroundColor Cyan
                                                    $OldVhds | Out-TableString
                                                    Invoke-Command -ComputerName $PSItem -ScriptBlock { Remove-Item -Path $using:OldVhds.FullName -Verbose } -Credential $ADCredential
                                                }
                                            }
                                            else {
                                                Write-Host "[$PSItem] No vhds found for deletion" -ForegroundColor Cyan
                                            }
                                        }
                                        catch {
                                            Write-Warning $_.Exception.Message
                                        }
                                        Clear-Variable -Name OldVhds, VhdTotalSize -ErrorAction Ignore
                                    }
                                    else {
                                        Write-Warning "[$PSItem] Connection failed"
                                    }
                                } # $FileServer | ForEach-Object {
                            } # elseif ($SiteProfilesSelection.Name -like "*-fs*") {

                            # Get disk stats if there were user profiles to delete
                            if ($UserProfileDeletion -or $VhdDeletion) {
                                Write-Host "`r`nPerformed Action: #43p - Deleted old user profiles" -ForegroundColor Cyan
                                Get-DiskUsage -ComputerName $SiteProfilesSelection.Name -Credential $ADCredential | Out-TableString
                            }
                            else {
                                Write-Host ''
                            }
                            Clear-Variable -Name SiteProfilesSelection, UserProfileDeletion, VhdDeletion -ErrorAction Ignore
                        }
                        '50' {
                            # #50
                            # Clear Print Queue
                            if (Get-Module -Name ActiveDirectory -ListAvailable) {
                                Clear-Variable -Name SiteServers -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#50 Select Remote Desktop Server(s) to clear print queue - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                                $SiteServers | foreach-object {
                                    Invoke-Command -ComputerName $PSItem.Name -ScriptBlock {
                                        Get-Service -Name Spooler | Stop-Service -Verbose
                                        Get-ChildItem -Path "C:\Windows\System32\spool\PRINTERS" | Remove-Item -Force -Verbose
                                        Get-Service -Name Spooler | Start-Service -Verbose
                                        Get-Printer | Select-Object -Property Name, PrinterStatus, JobCount | Sort-Object -Property Name
                                    } | Select-Object -Property PSComputerName, Name, PrinterStatus, JobCount | Format-Table -AutoSize | Out-String
                                }
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '51' {
                            # #51                            
                            # Printer list
                            if (Get-Module -Name ActiveDirectory -ListAvailable) {
                                Clear-Variable -Name SiteServers, PrinterSelection -ErrorAction Ignore
                                $SiteServers = Select-VcaSite -AU $Computername -Title "#51 Select Remote Desktop Server to show printer list - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single
                                if ($SiteServers) {
                                    $PrinterSelection = Invoke-Command -ComputerName $SiteServers.Name -ScriptBlock {
                                        $HostAddresses = @{}
                                        Get-WmiObject -ClassName Win32_TCPIPPrinterPort | ForEach-Object { $HostAddresses.Add($_.Name, $_.HostAddress) }
                                        Get-WmiObject -ClassName Win32_Printer | ForEach-Object {
                                            [pscustomobject] @{
                                                Name        = $_.Name
                                                DriverName  = $_.DriverName
                                                HostAddress = $HostAddresses[$_.PortName]
                                                Server      = $_.SystemName
                                                Location    = $_.Location
                                                PortName    = $_.PortName
                                            }
                                        }
                                    } | Sort-Object -Property Name |
                                        Select-Object -Property PSComputername, Name, DriverName, HostAddress, Server, Location, PortName |
                                        Out-GridView -Title "#51 [$($SiteServers.Name)] Select Printer(s) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple

                                    if ($PrinterSelection) {
                                        $PrinterSelection | foreach-object {
                                            if ($PSItem.HostAddress) {
                                                Start-Process "https://$($PSItem.HostAddress)"
                                            }
                                        }
                                    }
                                }
                            }
                            else {
                                Write-Warning 'ActiveDirectory module not found.'
                                Write-Warning 'For enhanced functionality please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
                            }
                        }
                        '60' { 
                            # #60
                            # Telecom File Templates
                            Start-Process -FilePath "$PSScriptRoot\Private\csv\Telecom\"
                        }
                        '60carr' {
                            # #60carr
                            Start-Process -FilePath "$PSScriptRoot\Private\csv\Telecom\primarycarriersecodaycarrier.xlsx"
                        }
                        '60con' {
                            # #60conn    
                            # Telecom Contacts
                            Import-Excel -Path "$PSScriptRoot\Private\csv\Telecom\VENDOR_VCA_IT_TEL_OPS_TEAM_01_2019_YM_REV02.xlsx" | Out-String
                        }
                        '60it' {
                            Start-Process -FilePath "$PSScriptRoot\Private\csv\Telecom\VCA-AT Installation Info.xlsx"
                        }
                        '60pbx' {
                            # #60pbx
                            Write-Host "PBX Information (primarycarriersecodaycarrier.xlsx):" -ForegroundColor Yellow
                            if (-not $PrimaryCarrier) {
                                $PrimaryCarrier = Import-Excel -Path "$PSScriptRoot\Private\csv\Telecom\primarycarriersecodaycarrier.xlsx" -WorksheetName 'SV9100 candidate sites'
                            }
                            if ($PrimaryCarrier) {
                                $PrimaryCarrierInfo = $PrimaryCarrier.Where( {
                                        $PSItem.'Candidate site - existing SV9100' -match "$(Convert-VcaAU -AU $ComputerName -Strip)" } )

                                if ($PrimaryCarrierInfo) {
                                    Write-Output $PrimaryCarrierInfo | Format-List | Out-String
                                }
                            }
                            Write-Host "UPS Information (VCA UPS List 11-18-16.xlsx):" -ForegroundColor Yellow
                            if (-not $TelecomUPS) {
                                $TelecomUPS = Import-Excel -Path "$PSScriptRoot\Private\csv\Telecom\VCA UPS List 11-18-16.xlsx"
                            }
                            if ($TelecomUPS) {
                                $TelecomUPSInfo = $TelecomUPS.Where( {
                                        $PSItem.AU -match "$(Convert-VcaAU -AU $ComputerName -Strip)" } )

                                if ($TelecomUPSInfo) {
                                    Write-Output $TelecomUPSInfo | Format-List | Out-String
                                }
                            }
                        }
                        '60sow' {
                            # #60sow
                            Start-Process -FilePath "$PSScriptRoot\Private\csv\Telecom\SOW_TEMPLATE.dotx"
                        }
                        '60toll' {
                            # #60toll
                            Start-Process -FilePath "$PSScriptRoot\Private\csv\Telecom\VCA ANTECH-Toll Free inventory Report_MCN_JH4109_with_dept_info.xlsx"
                        }
                        '61' {
                            # #61
                            Start-Process -FilePath "https://app.smartsheet.com/b/publish?EQBCT=08449f5e03f346d58c8fe2e0f1302064"
                        }
                        '62' {
                            # #62 Launch NEC Management Page
                            #if ($ComputerName -match '-vm$|-vm\d{1,2}') {
                            #if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$') {
                            if ($ComputerName -notmatch '-vm[\d]+') { # Don't know if this will this work with IPs or Stand-Alone hostnames.
                                # [Debug] # Write-Host "IFFalse ComputerName: "$ComputerName
                                # Non-Clustered Hostname?
                                $URLName = $ComputerName.Replace('-vm','-pbx')
                                # [Debug] # Write-Host "IFTrue ComputerName: "$URLName
                                Start-Process "https://$URLName"
                                Clear-Variable -Name URLName -ErrorAction Ignore
                            }
                            else { # Clustered Hostname?
                                # [Debug] # Write-Host "IFFalse ComputerName: "$ComputerName
                                $URLName = $ComputerName -replace ('-vm[\d]+','-pbx') # https://community.glideapps.com/t/regex-to-remove-number-from-string/36310 Post "Darren_Murphy Glide Certified Expert Dec 2021"
                                # [Debug] # Write-Host "IFTrue ComputerName: "$URLName
                                Start-Process "https://$URLName"
                                Clear-Variable -Name URLName -ErrorAction Ignore
                            }
                        }
                        '62u' {
                            # #62u Launch PBX UPS Management page
                            #if ($ComputerName -match '-vm$|-vm\d{1,2}') {
                            # if ($ComputerName -notmatch '^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$') {
                            if ($ComputerName -notmatch '-vm[\d]+') { # Don't know if this will this work with IPs or Stand-Alone hostnames.
                                # Non-Clustered Hostname?
                                # [Debug] # Write-Host "IFFalse ComputerName: "$ComputerName
                                $URLName = $ComputerName.Replace('-vm','-pbxUps')
                                # [Debug] # Write-Host "IFTrue ComputerName: "$URLName
                                Start-Process "http://$URLName"
                                Clear-Variable URLName
                            }
                            else { # Clustered Hostname?
                                # [Debug] # Write-Host "IFFalse ComputerName: "$ComputerName
                                $URLName = $ComputerName -replace ('-vm[\d]+','-pbxUps') # https://community.glideapps.com/t/regex-to-remove-number-from-string/36310 Post "Darren_Murphy Glide Certified Expert Dec 2021"
                                # [Debug] # Write-Host "IFTrue ComputerName: "$URLName
                                Start-Process "http://$URLName"
                                Clear-Variable URLName
                            }
                        }
                        '70' {
                            # #70
                            # Check shutdown events in system event logs
                            Clear-Variable -Name ServerSelection, ParallelTask -ErrorAction Ignore
                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                Out-GridView -Title "#70 Select Server(s) to check shutdown events- v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerSelection | Out-Null

                            if ($ServerSelection) {
                               $ScriptBlock = {
                                    if ($PSVersionTable.PSVersion.Major -ge 3) {
# To be removed in a future update after v.20251010                                        #Get-EventLog -LogName System -Source User32 -Newest 200 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
# To be removed in a future update after v.20251010                                        #(Get-EventLog -LogName System -Source User32, EventLog -Newest 200).Where({($_.EventID -eq '6008' -or $_.EventID -eq '6005')}) | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
# Leave this row for now.  The entry is Kinda OK I think to display this entry in the Ops Portal Tool output, an Admin account needs to launch the tool                                        #(Get-EventLog -LogName System -Source "Microsoft-Windows-Kernel-Boot" -InstanceId 20 -Newest 200).Where({($_.EventID -eq '20')}) | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
                                        (Get-EventLog -LogName System -Source EventLog, User32 -Newest 200).Where({($_.EventID -eq '1074' -or $_.EventID -eq '6005' -or $_.EventID -eq '6006' -or $_.EventID -eq '6008')}) | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
                                    }
                                    else {
# To be removed in a future update after v.20251010                                        #Get-EventLog -LogName System -Source User32 -Newest 200 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
                                        (Get-EventLog -LogName System -Source User32, EventLog -Newest 200) | Where-Object {($_.EventID -eq '1074' -or $_.EventID -eq '6005' -or $_.EventID -eq '6006' -or $_.EventID -eq '6008')} | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message, UserName
                                    }
                                }
                                $ParallelTask = Invoke-Command -ComputerName $ServerSelection.Name -ScriptBlock $ScriptBlock -AsJob -Credential $ADCredential
                                $ParallelTask | Wait-Job | foreach-object {
                                    try {
                                        $PSItem | Receive-Job -ErrorAction SilentlyContinue | Select-Object -Property PSCOmputerName, Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.'TimeWritten(Local)' -as [datetime]}}, Source, EventID, Message, UserName |
                                            Sort-Object -Property Timewritten, Index -Descending
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    finally {
                                        if ($ParallelTask) { $ParallelTask | Remove-Job -Force }
                                    }
                                } | Out-GridView -Title "#70 Shutdown Event Log (Newest 200) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                            }
                        }

                        '70a' {
                            # #70a
                            # Check events in application event logs
                            Clear-Variable -Name ServerSelection, ParallelTask -ErrorAction Ignore
                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                Out-GridView -Title "#70a Select Server(s) to check application events- v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerSelection | Out-Null

                            if ($ServerSelection) {
                               $ScriptBlock = {
                                    if ($PSVersionTable.PSVersion.Major -ge 3) {
                                        Get-EventLog -LogName application -Newest 100 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message
                                    }
                                    else {
                                        Get-EventLog -LogName application -Newest 100 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message
                                    }
                                }
                                $ParallelTask = Invoke-Command -ComputerName $ServerSelection.Name -ScriptBlock $ScriptBlock -AsJob -Credential $ADCredential
                                $ParallelTask | Wait-Job | foreach-object {
                                    try {
                                        $PSItem | Receive-Job -ErrorAction SilentlyContinue | Select-Object -Property PSCOmputerName, Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.'TimeWritten(Local)' -as [datetime]}}, Source, EventID, Message, UserName |
                                            Sort-Object -Property Timewritten, Index -Descending
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    finally {
                                        if ($ParallelTask) { $ParallelTask | Remove-Job -Force }
                                    }
                                } | Out-GridView -Title "#70a Application Event Log (Newest 100)- v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                            }
                        }
                        '70d' {
                            # #70d
                             # Check events in system event logs for disk errors
                             Clear-Variable -Name ServerSelection, ParallelTask -ErrorAction Ignore
                             $TitleTime = Get-Date -Format "yyyy-MM-dd hh:mm"
                             $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                             Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                 Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                 Out-GridView -Title "#70d Select Server(s) to check system events- v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerSelection | Out-Null
 
                             if ($ServerSelection) {
                                $ScriptBlock = {
                                    Get-EventLog -LogName System -Source disk -Newest 300  | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message
                                 }
                                 $ParallelTask = Invoke-Command -ComputerName $ServerSelection.Name -ScriptBlock $ScriptBlock -AsJob -Credential $ADCredential
                                 $ParallelTask | Wait-Job | foreach-object {
                                     try {
                                         $PSItem | Receive-Job -ErrorAction SilentlyContinue | Select-Object -Property PSCOmputerName, Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.'TimeWritten(Local)' -as [datetime]}}, Source, EventID, Message, UserName |
                                             Sort-Object -Property Timewritten, Index -Descending
                                     }
                                     catch {
                                         Write-Warning $_.Exception.Message
                                     }
                                     finally {
                                         if ($ParallelTask) { $ParallelTask | Remove-Job -Force }
                                     }
                                 } | Out-GridView -Title "#70d Windows System Log Disk Errors (Newest 300) - $TitleTime - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                             }
                        }
                        '70s' {
                            # #70s
                            # Check events in application event logs
                            Clear-Variable -Name ServerSelection, ParallelTask -ErrorAction Ignore
                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                Out-GridView -Title "#70s Select Server(s) to check system  events - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable ServerSelection | Out-Null

                            if ($ServerSelection) {
                               $ScriptBlock = {
                                    if ($PSVersionTable.PSVersion.Major -ge 3) {
                                        Get-EventLog -LogName System -Newest 100 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message
                                    }
                                    else {
                                        Get-EventLog -LogName System -Newest 100 | Select-Object -Property Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.TimeWritten|Out-String}}, Source, EventID, Message
                                    }
                                }
                                $ParallelTask = Invoke-Command -ComputerName $ServerSelection.Name -ScriptBlock $ScriptBlock -AsJob -Credential $ADCredential
                                $ParallelTask | Wait-Job | foreach-object {
                                    try {
                                        $PSItem | Receive-Job -ErrorAction SilentlyContinue | Select-Object -Property PSCOmputerName, Index, EntryType, TimeWritten, @{n='TimeWritten(Local)';e={$_.'TimeWritten(Local)' -as [datetime]}}, Source, EventID, Message, UserName |
                                            Sort-Object -Property  Timewritten, Index -Descending
                                    }
                                    catch {
                                        Write-Warning $_.Exception.Message
                                    }
                                    finally {
                                        if ($ParallelTask) { $ParallelTask | Remove-Job -Force }
                                    }
                                } | Out-GridView -Title "#70s System Event Log (Newest 100) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
                            }
                        }
                        '70w'{
                             # #70w Check woofware events in application event logs
                            Clear-Variable -Name ServerSelection, ParallelTask -ErrorAction Ignore
                            $TitleTime = Get-Date -Format "yyyy-MM-dd HH:mm"
                            $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
                                Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
                                Out-GridView -Title 'Select Server(s) to check system events' -OutputMode Multiple -OutVariable ServerSelection | Out-Null

                            if ($ServerSelection) {
                               #$ScriptBlock = {

                            $results = Invoke-Command -ComputerName $ServerSelection.Name {
                                $events = @()

                                    # Woofware & related
                                    $events += Get-WinEvent -FilterHashtable @{
                                        LogName = 'Application'
                                        ProviderName = 'Woofware'
                                    } -MaxEvents 200

                                    # Application Hang (Event ID 1002)
                                    $events += Get-WinEvent -ErrorAction SilentlyContinue -FilterHashtable @{
                                        LogName = 'Application'
                                        ProviderName = 'Application Hang'
                                        Id = 1002
                                    } -MaxEvents 200

                                $events | ForEach-Object {
                                    $msg = $_.Message

                                    # Extract structured fields from message text
                                    $threadIdentity  = if ($msg -match 'ThreadIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
                                    $windowsIdentity = if ($msg -match 'WindowsIdentity\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
                                    $machineName     = if ($msg -match 'MachineName\s*:\s*(.+)') { $matches[1].Trim() } else { 'N/A' }
                                    $exceptionType   = if ($msg -match 'Type\s*:\s*([^,]+)') { $matches[1].Trim() } else { 'N/A' }
                                    $MsgError        = if ($msg -match '(?s)Message\s*:\s*(AUID\s*=\s*.*?)(?:\r?\n\S|\Z)') {
                                                    $matches[1].Trim()
                                                } else {
                                                    'N/A'
                                                }
                                    #$msgPreview = if ($msg.Length -gt 300) { $msg.Substring(0,300) + ' ...' } else { $msg }

                                    [PSCustomObject]@{
                                        PSComputerName     = $env:COMPUTERNAME
                                        TimeCreated        = $_.TimeCreated
                                        EventID            = $_.Id
                                        RecordID           = $_.RecordId 
                                        MachineName        = $machineName
                                        ThreadIdentity     = $threadIdentity
                                        WindowsIdentity    = $windowsIdentity
                                        ExceptionType      = $exceptionType
                                        MessagError        = $MsgError
                                        FullMessage        = $msg   # Keep full message in case you want to display it later
                                    }
                                }
                            }

                            # Let user select events interactively
                            $selected = $results | Sort-Object TimeCreated, RecordID -Descending |
                                Out-GridView -Title "Select Woofware Events (Press OK to output) - $TitleTime" -PassThru

                            # Display selected event(s) in terminal
                            if ($selected) {
                                $selected | Format-List MachineName,TimeCreated, EventID, RecordID, HandlingInstanceID, ThreadIdentity, WindowsIdentity, ExceptionType, FullMessage
                                   }
                                               
                            }   
                        }
                        '70u' {
                            # #70u
                            Clear-Variable -Name WinUpdateLogResults -ErrorAction Ignore
                            $SiteServers = Select-VcaSite -AU $Computername -Title "#70u Select Server(s) to find Windows Update Initiator Account - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            if (-not $SiteServers) { break }

                            $WinUpdateLogResults = Find-WindowsUpdateInitiator -ComputerName $SiteServers.Name -Credential $ADCredential
                            if ($WinUpdateLogResults) {
                                $WinUpdateLogResults | Out-GridView -Title "#70u Select entries to send to console - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -PassThru | Out-TableString
                            }
                            else {
                                Write-Host "A windows update initiator account could not be found.`r`nPlease review the log at C:\temp\WindowsUpdate.log on the remote machine.`r`n" -ForegroundColor Cyan
                            }
                        }
                        '71d' {
                            # #71d
                            # Clean DMP and windows Temp on NS
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Invoke-command -computername $NetServices {
                                    Get-ChildItem -path "C:\*.dmp" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -confirm:$true
                                    Get-ChildItem -path "C:\windows\temp\*.*" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -confirm:$true
                                }
                            }
                        }
                        '71e' {
                            # #71e
                            ## Purge folders -- Engineering recommended
                            if (-not $NetServices) { break }
                            if ($NetServices) {
                                Invoke-Command -ComputerName $NetServices {
                                    $TSWorker = Get-ChildItem -path  "C:\Users\*\AppData\Roaming\Microsoft\Teams\Service Worker\CacheStorage\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="TSWorkerMB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $CDumps = Get-ChildItem -path  "C:\users\*\appdata\local\crashdumps\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="CrashDumpMB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $DaysToDelete = 7

                                    $W3SVC2 = Get-ChildItem -path  "C:\Inetpub\Log\W3SVC2\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Where-Object { ($_.CreationTime -lt $(Get-Date).AddDays(-$DaysToDelete))}|
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="W3SVC2MB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $W3SVC3 = Get-ChildItem -path  "C:\Inetpub\Log\W3SVC3\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Where-Object { ($_.CreationTime -lt $(Get-Date).AddDays(-$DaysToDelete))}|
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="W3SVC3MB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $VSTAGENT = Get-ChildItem -path  "C:\vstsagent\*\_diag\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="VSTAGENTMB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $AZAGENT = Get-ChildItem -path  "c:\azureagent\*\_Diag\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="AzureAGENTMB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    $ConfigNS = Get-ChildItem -path  "C:\config_ns_vm\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="ConfigNSMB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    Write-Host "`r`nData eligible for purge:"

                                    [PSCustomObject]@{
                                        ComputerName = $env:COMPUTERNAME
                                        TSWorkerMB   = $TSWorker.TSWorkerMB
                                        CDumpsMB     = $CDumps.CrashDumpMB
                                        W3SVC2MB     = $W3SVC2.W3SVC2MB
                                        W3SVC3MB     = $W3SVC3.W3SVC3MB
                                        VSTAGENTMB   = $VSTAGENT.VSTAGENTMB
                                        AZAGENTMB    = $AZAGENT.AzureAGENTMB
                                        ConfigNSMB   = $ConfigNS.ConfigNSMB
                                    } | Out-String
                                }

                                Invoke-Command -Computer $Netservices -ScriptBlock {
                                    Get-ChildItem -Path "C:\Users\*\AppData\Roaming\Microsoft\Teams\Service Worker\CacheStorage\*" | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\Users\*\AppData\Local\Crashdumps\*" | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\Inetpub\Log\W3SVC2" -Recurse -File | Where-Object CreationTime -LT (Get-Date).AddDays(-7) | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\Inetpub\Log\W3SVC3" -Recurse -File | Where-Object CreationTime -LT (Get-Date).AddDays(-7) | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\vstsagent\*\_diag\*", "C:\azureagent\*\_Diag\*" -Recurse | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\Windows\Temp\*" -Recurse | Remove-Item -Recurse -Force
                                    Get-Item -Path "C:\config-ns-vm" | Remove-Item -Recurse -Force
                                }
                            }
                        }
                        '71s' {
                            # #71s
                            ## Purge smpacs folders -- Engineering recommended
                            if (-not $NetServices) { break }
                            $SmpacsServer = $(Convert-VcaAU -AU $ComputerName -Suffix '-smpacs')
                            if ($SmpacsServer | Get-PingStatus) {
                                Invoke-Command -ComputerName $SmpacsServer {

                                    $W3SVC1 = Get-ChildItem -path  "C:\inetpub\logs\LogFiles\W3SVC1\*" -Recurse -Force -ErrorAction SilentlyContinue |
                                        Where-Object { ($_.CreationTime -lt $(Get-Date).AddDays(-2))}|
                                        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue |
                                        Select-Object -Property @{n="W3SVC1MB"; e={[math]::Round($_.Sum/1MB,2)}}

                                    Write-Host "`r`nData eligible for purge:" -ForegroundColor Cyan

                                    [PSCustomObject]@{
                                        ComputerName = $env:COMPUTERNAME
                                        W3SVC1MB     = $W3SVC1.W3SVC1MB
                                    } | Format-Table -AutoSize | Out-String

                                    Get-ChildItem -Path "C:\inetpub\logs\LogFiles\W3SVC1" -Recurse -File | Where-Object CreationTime -LT (Get-Date).AddDays(-2) | Remove-Item -Recurse -Force
                                    Get-ChildItem -Path "C:\Windows\Temp\*" -Recurse | Remove-Item -Recurse -Force
                                }

                                Write-Host "`r`nPerformed Action: #71s - Deleted temps" -ForegroundColor Cyan
                                Get-DiskUsage -ComputerName $SmpacsServer -Credential $ADCredential | Out-TableString
                            }
                        }
                        '80' {
                            # #80
                            # Clear WOOFware Global Medical Record (GMR) Cache
                            Clear-Variable -Name WhatUsers -ErrorAction Ignore
                            if ((-not $Cluster) -and $NetServices) {
                                if ($NetServices | Get-Pingstatus) { $WhatUsers = whatusers -ComputerName $NetServices }
                            }
                            elseif ($Cluster) {
                                Clear-Variable -Name NetservicesResults -ErrorAction Ignore
                                $NetservicesResults = $(
                                    $NetServices | ForEach-Object {
                                        if ($PSItem | Get-Pingstatus) { $PSItem }
                                    }
                                )
                                $WhatUsers = whatusers -ComputerName $NetServicesResults
                            }

                            if ($WhatUsers) {
                                Clear-Variable -Name WhatUsersSelection -ErrorAction Ignore
                                $WhatUsers | Out-GridView -Title "#80 $SiteAUNumber - Total logged on users: $(($WhatUsers | Where-Object UserName -ne '').Count) - Select user to Clear GMR Cache - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable WhatUsersSelection | Out-String
                                if ($WhatUsersSelection.Username) {
                                    Remove-Item -Path "\\$($WhatUsersSelection.Computer)\c$\Users\$($WhatUsersSelection.Username)\AppData\Local\woofware_msal_cache.bin" -Verbose -Confirm
                                }
                            }
                        }
                        '81' {
                            # #81
                            # Woofware Reports Website
                            $ComputerNameStripped = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Start-Process "http://$ComputerNameStripped-db/reports/browse/WOOFware%20Reports"
                        }
                        '82' {
                            # #82
                            # Fuse Website
                            $ComputerNameStripped = Convert-VcaAu -AU $ComputerName -Suffix ''
                            Start-Process "https://$ComputerNameStripped-fuse:8443"
                        }
                        '83' {
                            # #83
                            # Restart Sparky Services
                            if (Get-Module -Name ActiveDirectory -ListAvailable) {
                                Clear-Variable -Name SiteServers, SiteAU -ErrorAction Ignore
                                $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
                                Get-ADComputer -Filter "Name -like '$SiteAU-ns*' -and OperatingSystem -like '*Server*'" -Properties IPv4Address, OperatingSystem |
                                    Select-Object Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-Pingstatus } } | Sort-Object Name |
                                    Out-GridView -Title "#83 Select Remote Desktop Server to Reset Sparky Services - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable SiteServers | Out-Null
                                $SiteServers | foreach-object {
                                    Invoke-command -ComputerName $PSItem.Name { Get-Service -Name Sparky* | Restart-Service -Verbose }
                                }
                            }
                        }
                        '98' {
                            # #98
                            $Contacts2 = Import-Csv -Path "$PSScriptRoot\Private\csv\Contacts.csv"
                            $Contacts2 | Format-Table -AutoSize -Wrap | Out-String
                            #Write-Host ''
                        }
                        '98l' {
                            # #98L
                            # ServiceNow KB Article
                            Start-Process "https://marsvh.service-now.com/kb_view.do?sysparm_article=KB0010651"
                            Write-Host ''
                        }
                        '99' {
                            # #99
                            # Clear credentials
                            Clear-Variable -Name TargetCredsDeletion, TargetCreds_Item -ErrorAction Ignore
                            $TargetCreds = @(
                                'vcahospesxi'
                                'vcahospilo'
                                'vcadomaincreds'
                                'vcasnowapi'
                                'vcaemailcreds'
                            )
                            $CredentialTargetsObj = $TargetCreds | foreach-object {
                                $TargetCreds_Item = $PSItem
                                Get-StoredCredential -Target $PSItem | Select-Object -Property UserName, Password, @{n='Target';e={$TargetCreds_Item}}
                            }
                            @($CredentialTargetsObj) + [pscustomobject]@{Username = "----- Refresh cached passwords -----"; Password = "Useful when password has been updated in a separate session"} +
                                [pscustomobject]@{Username = "----- Update Admin Credentials -----"; Password = ""} |
                                Out-GridView -Title "#99 Select Credential(s) to delete - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable TargetCredsDeletion | Out-Null

                            if ($TargetCredsDeletion) {
                                if ($TargetCredsDeletion.UserName -like "*Refresh cached passwords*") {
                                    $EsxiCredential = Get-StoredCredential -Target vcahospesxi
                                    $IloCredential = Get-StoredCredential -Target vcahospilo
                                    $ADCredential = Get-StoredCredential -Target vcadomaincreds
                                    $SNOWAPICredential = Get-StoredCredential -Target vcasnowapi
                                    $EmailCredential = Get-StoredCredential -Target vcaemailcreds
                                    Write-Host "`r`nCached passwords refreshed`r`n" -ForegroundColor Cyan
                                }
                                elseif ($TargetCredsDeletion.UserName -like "*Update Admin Credentials*") {
                                    $script:ADCredential = Get-ADCreds
                                    if ($ADCredential) {
                                        New-StoredCredential -Credentials $script:ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                        Write-Host "`r`nNew Admin Credentials set`r`n" -ForegroundColor Cyan
                                    }
                                }
                                else {
                                    $TargetCredsDeletion | foreach-object {
                                        Write-Host "[$($PSItem.Target)] Credential to be deleted" -ForegroundColor Cyan
                                        Remove-StoredCredential -Target $PSItem.Target -Verbose

                                        $TargetCredsVariable = switch ($PSItem.Target) {
                                            'vcahospesxi' { 'EsxiCredential' }
                                            'vcahospilo' { 'IloCredential' }
                                            'vcadomaincreds' { 'ADCredential' }
                                            'vcasnowapi' { 'SNOWAPICredential' }
                                            'vcaemailcreds' { 'EmailCredential' }
                                        }
                                        Clear-Variable -Name $TargetCredsVariable -ErrorAction Ignore
                                        Write-Host ''
                                    }
                                }
                            }
                        }
                        '99r' {
                            # #99r
                            Invoke-CachedPwRefresh
                        }
                        '99u' {
                            # #99u
                            Clear-Variable -Name NewADCredential -ErrorAction Ignore
                            $NewADCredential = Get-ADCreds
                            if ($NewADCredential) {
                                $script:ADCredential = $NewADCredential
                                New-StoredCredential -Credentials $script:ADCredential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
                                Write-Host "`r`nNew Admin Credentials set`r`n" -ForegroundColor Cyan
                            }
                        }
                        '99L' {
                            # #99L
                            $LapsSelection = Select-VcaSite -AU $Computername -Title "#99L Select server(s) to retrieve Local Admin password - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple
                            if ($LapsSelection.Name) {
                                $LapsSelection.Name | foreach-object {
                                    (Get-AdmPwdPassword -ComputerName $_ | Format-List | Out-String) -replace '\r\n\r\n\r\n', ''
                                }
                            }
                        }
                        '999' {
                            # #999
                            # new session
                            #Start-Process -FilePath "$PSScriptRoot\VCA Ops Portal.cmd" -WorkingDirectory $PSScriptRoot
                            Start-VcaOpsPortal
                        }
                        default {
                            #meh
                        }
                    } #switch
                    if ($ComputerName) {
                        # Invoke new action prompt
                        Clear-Variable -Name MenuAction
                    }
                } #else
            } #while (-not $MenuAction)
        } #if ($ComputerName -match '^[a-zA-Z0-9-.]+$|^\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b$')
        elseif ($ComputerName -eq '999') {
            # new session
            Start-VcaOpsPortal
            Clear-Variable -Name ComputerName
        }
        elseif ($ComputerName) {
            # Invoke hostname change prompt
            Clear-Variable -Name ComputerName
        }
    } #while (-not $ComputerName)
} #function


# FUNCTIONS
function Read-Choice {
    param(
        [string]$Title = 'Yes or no?',
        [string]$Message = 'Proceed?',
        [int]$DefaultChoice = 0
    )
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "This means Yes"
    $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "This means No"
    $Options = [System.Management.Automation.Host.ChoiceDescription[]]($No, $Yes)

    Write-Output $Host.UI.PromptForChoice($Title, $Message, $Options, $DefaultChoice)
}



function Start-VcaOpsPortal {
    # check if running locally, or local copy exists then launch respective instance
    if ((Test-Path -Path "$env:USERPROFILE\Desktop\VCA Ops Portal\VCAHospLauncher.ps1") -or (-not ([System.Uri]$PSCommandPath).IsUnc)) {
        # if currently running session is not the standard local path
        if ($PSCommandPath -ne "$env:USERPROFILE\Desktop\VCA Ops Portal\VCAHospLauncher.ps1" -and (-not ([System.Uri]$PSCommandPath).IsUnc)) {
            Invoke-Expression -Command "cmd /c start powershell -NoExit -NoProfile -ExecutionPolicy Bypass -Command { . `"$PSCommandPath`" }"
        }
        else {
            #& "$env:USERPROFILE\Desktop\VCA Ops Portal\VCA Ops Portal.cmd"
            Invoke-Expression -Command "cmd /c start powershell -NoExit -NoProfile -ExecutionPolicy Bypass -Command { . `"$env:USERPROFILE\Desktop\VCA Ops Portal\VCAHospLauncher.ps1`" }"
        }
    }
    else {
        # launch network copy if local copy doesn't exist
        #Invoke-Expression -Command "cmd /c `"$PSScriptRoot\VCA Ops Portal.cmd`""
        Invoke-Expression -Command "cmd /c start powershell -NoExit -NoProfile -ExecutionPolicy Bypass -Command { . `"$PSCommandPath`" }"
        #& "$PSScriptRoot\VCA Ops Portal.cmd"
    }
}



function Get-VcaOpsPortalVersion {
    # #robo #ver #version
    param(
        [decimal]$Version,
        [switch]$OnlyNew,
        [switch]$ShowWarning
    )
    try {
# Check network share asynchronously so it doesn't indefinitely freeze portal tool in the case the network share is 'down'
$ps = [powershell]::Create().AddScript("
    Test-Path -Path '\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\Version.txt'
")
# execute it asynchronously
$handle = $ps.BeginInvoke()

# Wait 3000 milliseconds for it to finish
if(-not $handle.AsyncWaitHandle.WaitOne(3000)){
    if ($ShowWarning.IsPresent) { Write-Warning "Network share timed out, try again later." }
    return
}
# WaitOne() returned $true, let's fetch the result
$NetShareOnlineStatus = $ps.EndInvoke($handle)

        if ($NetShareOnlineStatus) {
            $VersionFile = Get-Content -Path '\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\Version.txt' -ErrorAction Stop
            # Newer version available
            if ([decimal]($VersionFile.trim()) -gt [decimal]$Version) {
                if (-not ([System.Uri]$PSCommandPath).IsUnc) {
                    # running locally
                    Write-Host "Version $VersionFile is now available, #robo is being automatically activated.  `nPlease answer 'Yes' to the following questions to upgrade Ops Portal Tool to version v.$($VersionFile)." -ForegroundColor Magenta
                    Invoke-VcaOpsPortalUpdate
                }
                else {
                    # running from network drive
                    Write-Host "Version $VersionFile is now available, use #999 action to launch updated session." -ForegroundColor Magenta
                    Write-Host "Use #robo action instead to install locally for improved performance." -ForegroundColor Magenta
                }
            }
            elseif (-not $OnlyNew.IsPresent -and ([decimal]($VersionFile.trim()) -eq [decimal]$Version)) {
                Write-Host "No update is available`r`n" -ForegroundColor Cyan
            }
            else {
                if (-not $OnlyNew.IsPresent) {
                    Write-Host "How are you on a newer version? :o`r`n" -ForegroundColor Yellow
                }
            }
            if (([System.Uri]$PSCommandPath).IsUnc) {
                Write-Host "You are running Ops Portal from the network share..." -ForegroundColor Magenta
                Write-Host "#robo is being automatically activated.  `nPlease answer 'Yes' to the following questions to install Ops Portal Tool locally and improve performance by running it on your local machine." -ForegroundColor Magenta
                Invoke-VcaOpsPortalUpdate
            }
        }
    }
    catch {
        # intentionally left blank
    }
}



function Invoke-VcaOpsPortalUpdate { #robo
    if ((Read-Choice -Title "#robo Copy VCA Ops Portal to $env:USERPROFILE\Desktop\VCA Ops Portal? - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -Message 'Copy to Desktop?' -DefaultChoice 1) -eq 1) {
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # /mir	Mirrors a directory tree (equivalent to /e plus /purge). Using this option with the /e option and a destination directory, overwrites the destination directory security settings.
        # /mt:<n>	Creates multi-threaded copies with n threads. n must be an integer between 1 and 128. The default value for n is 8. For better performance, redirect your output using /log option.
        # /xx	Excludes extra files and directories present in the destination but not the source. Excluding extra files will not delete files from the destination.
        # /xd <directory>[ ...]	Excludes directories that match the specified names and paths.
        # /r:<n>	Specifies the number of retries on failed copies. The default value of n is 1,000,000 (one million retries).
        # /w:<n>	Specifies the wait time between retries, in seconds. The default value of n is 30 (wait time 30 seconds).
        # /np	Specifies that the progress of the copying operation (the number of files or directories copied so far) will not be displayed.

        robocopy "\\vcaantech.com\folders\data2\Corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal" "$env:USERPROFILE\Desktop\VCA Ops Portal" /MIR /MT:16 /XX /XD Archive Reports /NFL /NDL /R:5 /W:5 /NP
        # if successful robocopy, prompt for launching new session
        if (Test-Path -Path "$env:USERPROFILE\Desktop\VCA Ops Portal\VCAHospLauncher.ps1") {
            # check if desktop shortcut exists, ask to create
            if (-not (Test-Path -Path "$env:USERPROFILE\Desktop\VCA Ops Portal.cmd.lnk")) {
                if ((Read-Choice -Title "VCA Ops Portal Tool desktop shortcut doesn't exist" -Message 'Create desktop shortcut?' -DefaultChoice 1) -eq 1) {
                    $WshShell = New-Object -ComObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\VCA Ops Portal.cmd.lnk")
                    $Shortcut.TargetPath = "$env:USERPROFILE\Desktop\VCA Ops Portal\VCA Ops Portal.cmd"
                    $Shortcut.Save()
                }
            }
            if ((Read-Choice -Title "Launch updated VCA Ops Portal Tool session?" -Message '(Current session will stay open)' -DefaultChoice 1) -eq 1) {
                Start-VcaOpsPortal
            }
        }
        else {
            Write-Host "Could not find `"$env:USERPROFILE\Desktop\VCA Ops Portal\VCAHospLauncher.ps1`"`r`n" -ForegroundColor Yellow
        }
    }
}



function Get-ADCreds {
    param(
        $UserName
    )
    $CredentialParams = @{}
    if ($ADCredential.UserName) {
        $CredentialParams.UserName = $ADCredential.UserName
    }
    $Credential = Get-Credential -Message 'Enter Admin Credentials' @CredentialParams
    if ($Credential.UserName -eq $null -or $Credential.GetNetworkCredential().Password -eq $null) { return }
    $ADUsername = $(if ($Credential.UserName -notmatch '\\') { "vcaantech\$($Credential.UserName)" } else { $Credential.UserName })
    $ADPassword = $Credential.GetNetworkCredential().Password

    if ($Credential) {
        # Get current domain using logged-on user's credentials
        $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $ADUsername, $ADPassword)

        if ($Domain.Name -eq $null) {
            Write-Warning "Authentication failed - please verify your username and password."
            return
        }

        $secpasswd = ConvertTo-SecureString $ADPassword -AsPlainText -Force
        $Credential = New-Object System.Management.Automation.PSCredential ($ADUsername, $secpasswd)

        Write-Output $Credential
    }
}



function Get-EsxiCredential {
    Get-Credential -UserName 'root' -Message 'Enter ESXi Host Credentials'
}



function Set-EsxiCredential {
    param(
        [parameter(Mandatory)]
        [pscredential]$Credential
    )
    if (-not (Get-StoredCredential -Target vcahospesxi)) {
        if ($Credential.GetNetworkCredential().Password) {
            New-StoredCredential -Credentials $Credential -Target vcahospesxi -Type Generic -Persist LocalMachine | Out-Null
        }
    }
}



function Enable-SSH {
    param(
        [parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName,
        [pscredential]$Credential = $(Get-Credential)
    )
    foreach ($ComputerName_Item in $ComputerName) {
        $VIServer = Connect-VIServer -Server $ComputerName -Credential $Credential -WarningAction SilentlyContinue
        if ($VIServer.IsConnected) {
            Write-Host "[$ComputerName] Checking SSH Service"

            $SSHStatus = Get-VMHostService -Refresh | Where-Object { $_.Label -eq 'SSH' }
            if ($SSHStatus.Running -ne $true) {
                $SSHStatus | Start-VMHostService | Format-Table -AutoSize | Out-String
            }
            else {
                Write-Warning "[$ComputerName] SSH Service is already running"
            }
            Write-Host ''
        }
    }
}



function Disable-SSH {
    param(
        [parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName,
        [pscredential]$Credential = $(Get-Credential)
    )
    foreach ($ComputerName_Item in $ComputerName) {
        $VIServer = Connect-VIServer -Server $ComputerName -Credential $Credential -WarningAction SilentlyContinue
        if ($VIServer.IsConnected) {
            Write-Host "[$ComputerName] Checking SSH Service"

            $SSHStatus = Get-VMHostService -Refresh | Where-Object { $_.Label -eq 'SSH' }
            if ($SSHStatus.Running -eq $true) {
                $SSHStatus | Stop-VMHostService -Confirm:$false | Format-Table -AutoSize | Out-String
            }
            else {
                Write-Warning "[$ComputerName] SSH Service is not running"
            }
            Write-Host ''
        }
    }
}



#13
# Needed to rename title bar of a running process
Add-Type -Namespace PInvoke -Name User -MemberDefinition @"
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern bool SetWindowText(IntPtr hwnd, String lpString);
"@
function Start-PingInfoView {
    # #13
    param(
        [parameter(Position = 0)]
        [string[]]$ComputerName,
        [string]$WorkingPath = "$PSScriptRoot\Private\bin"
    )
    # Retrieve user local/source/vpn IP to display in titlebar
    if (Test-Path -Path "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe") {
        $vpn_stats = New-Object PSObject; &"C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe" stats |
            Where-Object {(($_ -match ':') -and ($_ -notlike "*>>*"))} |
            ForEach-Object {$Item = ($_.Trim() -split ': ').trim()
            $vpn_stats | Add-Member -MemberType NoteProperty -Name $($Item[0]) -Value $Item[1] -ErrorAction SilentlyContinue
    }
    if ($vpn_stats.'Connection State' -eq 'Connected') {
        $ConnectionInfo = "- [Source IP: $($vpn_stats.'Client Address (IPv4)') via ($([System.Net.Dns]::GetHostEntry($vpn_stats.'Server Address').HostName) - $($vpn_stats.'Connection State'))]"
    }
    } else {
        $ConnectionInfo = (Get-NetIPConfiguration).IPv4Address.IPAddress | Where-Object { $_ -notmatch "^192\.|^169\."} | ForEach-Object {
            "- [Source IP: $_ ($([System.Net.Dns]::GetHostEntry($_).HostName))]"
        }
        #$ConnectionInfo = "Source IP(s): $((Get-NetIPConfiguration).IPv4Address.IPAddress | Where-Object { $_ -notmatch "^192\.|^169\."})"
    }

    if ($ComputerName) {
        $ComputerName | Out-File -FilePath "$WorkingPath\Servers.txt"

        $PingInfoViewProcObj = Start-Process -FilePath "$WorkingPath\PingInfoView.exe" -ArgumentList '/loadfile Servers.txt /PingEvery 1 /PingEverySeconds 5' -WorkingDirectory $WorkingPath -PassThru

        # Rename PingInfoView title bar to include ops portal version number
        Start-Sleep -Milliseconds 500
        if ($Computername.count -gt 1) {
            $PingInfoViewProcObj | Where-Object MainWindowHandle |
            ForEach-Object { [PInvoke.User]::SetWindowText($PSItem.MainWindowHandle, "$($PSItem.MainWindowTitle) - VCA Ops Portal v.$Version - $(Convert-VcaAU -AU (@($ComputerName[0])) -Prefix AU -Suffix '' -NoLeadingZeros) $ConnectionInfo") } | Out-Null
        }
        elseif ($Computername.count -eq 1) {
            $PingInfoViewProcObj | Where-Object MainWindowHandle |
            ForEach-Object { [PInvoke.User]::SetWindowText($PSItem.MainWindowHandle, "$($PSItem.MainWindowTitle) - VCA Ops Portal v.$Version - $ComputerName $ConnectionInfo") } | Out-Null
        }
    }
    else {
        $PingInfoViewProcObj = Start-Process -FilePath "$WorkingPath\PingInfoView.exe" -ArgumentList '/PingEvery 1 /PingEverySeconds 5' -PassThru

        # Rename PingInfoView title bar to include ops portal version number
        Start-Sleep -Milliseconds 500
        $PingInfoViewProcObj | Where-Object MainWindowHandle |
            ForEach-Object { [PInvoke.User]::SetWindowText($PSItem.MainWindowHandle, "$($PSItem.MainWindowTitle) - VCA Ops Portal v.$Version $ConnectionInfo") } | Out-Null
    }
}



filter Get-PingStatus {
    try {
        $ErrorActionPreference = 'Stop'
        $obj = New-Object system.Net.NetworkInformation.Ping
        if (($obj.Send($PSItem, '1000')).status -eq 'Success') { 'Online!' }
    }
    catch {
        #intentionally left blank
    }
    finally {
        $ErrorActionPreference = 'Continue'
    }
}



function Get-UserMemory {
    # #11m
    param(
        [string[]]$NetServices
    )
    # user session memory
    if (-not $NetServices) { break }
    if ($NetServices) {

        Clear-Variable -Name Processes, ProcessesFormatted -ErrorAction Ignore

        $Processes = Invoke-Command -ComputerName $Netservices {
            Get-Process -IncludeUserName | Select-Object -Property Username, WorkingSet64
        }

        if ($Processes) {
            $ProcessesFormatted = $Processes.Where( { $_.UserName -like 'VCAANTECH*' }) | Group-Object -Property Username |
                Select-Object -Property Name, @{n = 'MemoryMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
                Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "MemoryMB"; Descending = $True }

            $ProcessesFormatted | Out-GridView -Title "#11m Memory usage by user - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" | Out-Null
            $ProcessesFormatted | Format-Table -Property Name, @{n = 'MemoryMB'; e = { '{0:N2}' -f $_.'MemoryMB' } }, PSComputerName

            $Processes.Where( { $_.UserName -like 'VCAANTECH*' }) | Group-Object -Property PSComputerName |
                Select-Object -Property Name, @{n = 'MemoryMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
                Sort-Object -Property 'PSComputername' | Format-Table -Property @{n = 'MemoryMB'; e = { '{0:N2}' -f $_.'MemoryMB' } }, PSComputerName
        }
    }
}



function Get-HeavyMem {
    # #11h
    param(
        [string[]]$ComputerName
    )
    # Memory sum by user for HeavyHitters Teams, Chrome, Woofware, Runtimebroker
    if (-not $NetServices) { break }
    if ($NetServices) {

        Clear-Variable -Name Processes -ErrorAction Ignore

        $Processes = Invoke-Command -ComputerName $ComputerName {
            Get-Process -name teams, ms-teams, chrome, runtimebroker, vca.Sparky.shell, msedgewebview2 -IncludeUserName -ErrorAction Ignore   |
            Select-Object -Property PSComputerName, Username, ProcessName, WorkingSet64, Id, Handle, Path
        }

        if ($Processes) {
            $Processes | Where-object Processname -eq 'teams' | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'TeamsMemMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "TeamsMemMB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'TeamsMemMB'; e = { '{0:N2}' -f $_.'TeamsMemMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'teams' | Group-Object -Property PSComputerName  |
            select-object @{n = 'TeamsTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'TeamsTotalMB'; e = { '{0:N2}' -f $_.'TeamsTotalMB' } }, PSComputerName -NoNewLine
            $Processes | Where-object Processname -eq 'ms-teams' | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'MS-TeamsMemMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "MS-TeamsMemMB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'MS-TeamsMemMB'; e = { '{0:N2}' -f $_.'MS-TeamsMemMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'ms-teams' | Group-Object -Property PSComputerName  |
            select-object @{n = 'MS-TeamsTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'MS-TeamsTotalMB'; e = { '{0:N2}' -f $_.'MS-TeamsTotalMB' } }, PSComputerName -NoNewLine
           
            write-host "`r`n*** msedgewebview2 is a secondary process for ms-teams application *** `r`n" -ForegroundColor Cyan
            $Processes | Where-object Processname -eq 'msedgewebview2'  | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'msedgewebview2MB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "msedgewebview2MB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'msedgewebview2MB'; e = { '{0:N2}' -f $_.'msedgewebview2MB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'msedgewebview2' | Group-Object -Property PSComputerName  |
            select-object @{n = 'msedgewebview2TotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'msedgewebview2TotalMB'; e = { '{0:N2}' -f $_.'msedgewebview2TotalMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'Chrome' | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'ChromeMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "ChromeMB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'ChromeMB'; e = { '{0:N2}' -f $_.'ChromeMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'Chrome' | Group-Object -Property PSComputerName  |
            select-object @{n = 'ChromeTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'ChromeTotalMB'; e = { '{0:N2}' -f $_.'ChromeTotalMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'vca.sparky.shell'  | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'WOOFwareMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "WOOFwareMB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'WOOFwareMB'; e = { '{0:N2}' -f $_.'WOOFwareMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'vca.sparky.shell' | Group-Object -Property PSComputerName  |
            select-object @{n = 'WOOFwareTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'WOOFwareTotalMB'; e = { '{0:N2}' -f $_.'WOOFwareTotalMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'RuntimeBroker'  | Group-Object -Property Username |
            select-object @{n='Username';e={$_.Name}}, @{n = 'RuntimeBrokerMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property @{Expression = "PSComputerName"; Descending = $False }, @{Expression = "RuntimeBrokerMB"; Descending = $True } |
            Out-TableString -Property Username, @{n = 'RuntimeBrokerMB'; e = { '{0:N2}' -f $_.'RuntimeBrokerMB' } }, PSComputerName -NoNewLine

            $Processes | Where-object Processname -eq 'RuntimeBroker' | Group-Object -Property PSComputerName  |
            select-object @{n = 'RuntimeBrokerTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'RuntimeBrokerTotalMB'; e = { '{0:N2}' -f $_.'RuntimeBrokerTotalMB' } }, PSComputerName -NoNewLine
        }
    }
}



function Get-HeavyMemTot {
    # #11ht
    param(
        [string[]]$NetServices
    )
    # Memory sum by user for HeavyHitters Teams, Chrome, Woofware, Runtimebroker
    if (-not $NetServices) { break }
    if ($NetServices) {

        Clear-Variable -Name Processes -ErrorAction Ignore

        $Processes = Invoke-Command -ComputerName $Netservices {
            Get-Process -name teams, ms-teams,chrome, runtimebroker, vca.Sparky.shell, msedgewebview2 -IncludeUserName -ErrorAction Ignore   |
            Select-Object -Property PSComputerName, Username, ProcessName, WorkingSet64, Id, Handle, Path
        }

        if ($Processes) {

            $Processes | Where-object Processname -eq 'teams' | Group-Object -Property PSComputerName  |
            select-object Name, @{n = 'TeamsTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property 'PSComputername' | Out-TableString -Property @{n = 'TeamsTotalMB'; e = { '{0:N2}' -f $_.'TeamsTotalMB' } }, PSComputerName

            $Processes | Where-object Processname -eq 'ms-teams' | Group-Object -Property PSComputerName  |
            select-object Name, @{n = 'MS-TeamsTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property 'PSComputername' | Out-TableString -Property @{n = 'MS-TeamsTotalMB'; e = { '{0:N2}' -f $_.'MS-TeamsTotalMB' } }, PSComputerName

            write-host "`r`n*** msedgewebview2 is a secondary process for ms-teams application *** `r`n" -ForegroundColor Cyan
            $Processes | Where-object Processname -eq 'msedgewebview2' | Group-Object -Property PSComputerName  |
            select-object @{n = 'msedgewebview2TotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n='PSComputerName';e={$_.Name}} |
            Sort-Object -Property PSComputerName | Out-TableString -Property @{n = 'msedgewebview2TotalMB'; e = { '{0:N2}' -f $_.'msedgewebview2TotalMB' } }, PSComputerName -NoNewLine
            
            $Processes | Where-object Processname -eq 'Chrome' | Group-Object -Property PSComputerName  |
            select-object Name, @{n = 'ChromeTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property 'PSComputername' | Out-TableString -Property @{n = 'ChromeTotalMB'; e = { '{0:N2}' -f $_.'ChromeTotalMB' } }, PSComputerName

            $Processes | Where-object Processname -eq 'vca.sparky.shell' | Group-Object -Property PSComputerName  |
            select-object Name, @{n = 'WOOFwareTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property 'PSComputername' | Out-TableString -Property @{n = 'WOOFwareTotalMB'; e = { '{0:N2}' -f $_.'WOOFwareTotalMB' } }, PSComputerName

            $Processes | Where-object Processname -eq 'RuntimeBroker' | Group-Object -Property PSComputerName  |
            select-object Name, @{n = 'RuntimeBrokerTotalMB'; e = { [decimal]('{0:N0}' -f (($_.Group | Measure-Object WorkingSet64 -Sum).Sum / 1MB)) } }, @{n = 'PSComputerName'; e = { @($_.Group.PSComputerName)[0] } } |
            Sort-Object -Property 'PSComputername' | Out-TableString -Property @{n = 'RuntimeBrokerTotalMB'; e = { '{0:N2}' -f $_.'RuntimeBrokerTotalMB' } }, PSComputerName
        }
    }
}



function Get-Quser {
    # #7
    param(
        [string[]]$ComputerName
    )
    if (-not $ComputerName) { break }

    $ComputerName | foreach-object {
        try {
            Clear-Variable -Name QuserCount -ErrorAction Ignore
            Write-Host "`r`nCMD or PS Command: quser /server:$PSItem`r`n" -ForegroundColor Cyan
            Write-Host "Server: $PSItem ($((Resolve-DnsName -Name $PSItem -ErrorAction Stop).IPAddress))`r`n"

            if ($PSItem | Get-PingStatus) {
                quser "/server:$PSItem" | Tee-Object -Variable QuserCount
                Write-Host "`r`nCount of users: $(@($QuserCount).count - 1)"
            }
            else {
                Write-Warning "[$PSItem] Connection failed"
            }
        }
        catch {
            #intentionally left blank
        }
        Write-Host ''
    }
}


function Get-VcaSiteHostname {
    [CmdletBinding()]
    param(
        [parameter(Mandatory)]
        [string[]]$ComputerName,
        [string[]]$Cluster,
        [string[]]$ClusterSite,
        [string[]]$NetServices
    )

    <# Error checking text to see how the variable change as we override $Cluster for vCenter Cluster sites that moved to the cloud and no longer have ESXi hosts.  We still want those hospitals to make it to the multi-server site logic.
        Write-Host "In function"
        Write-Host "ComputerName: $ComputerName"
        Write-Host "Cluster: $Cluster"
        Write-Host "ClusterSite: $ClusterSite"
        Write-Host "NetServices: $NetServices"
        Write-Host "Count of hosts in `$NetServices: $(@($Netservices).count)" # Error checking code. Count how many objects are in $NetServices which only contain NSes.
    #>

    if (-not $Cluster) { # As long as $Cluster is not empty
        if (@($Netservices).count -gt 1) { # Make sure that there's more than 1 host in the variable
            # Write-Host "I made it inside the if" # Error checking code
            $Cluster = @("Manually_look_for_hidden_ESXi_&_iLO_hosts") + $NetServices # Re-Write the variable so it goes into the multi-cluster selection below (when vCenter Clusters no longer exist). A <space> is needed at the end of the one that's being pre-pended as to not link it to the first obeject in the variable.  Exmaple: Manually_look_for_hidden_ESXi_&_iLO_hosts vs. Manually_look_for_hidden_ESXi_&_iLO_hostsh900-ns1
            # $Cluster = $NetServices # Use this line when all MSAs go away and the -san gets removed from the list below.  Don't forget to delete the line above. # Re-Write the variable so it goes into the multi-cluster selection below (when vCenter Clusters no longer exist). A <space> is needed at the end of the one that's being pre-pended as to not link it to the first obeject in the variable.  Exmaple: Manually_look_for_hidden_ESXi_&_iLO_hosts vs. Manually_look_for_hidden_ESXi_&_iLO_hostsh900-ns1
        }
    }

    $ComputerNameStripped = Convert-VcaAu -AU $ComputerName -Suffix ''
    if ((-not $Cluster) -and $NetServices) {
        $SiteHostnames = @(
            "$ComputerNameStripped-gw"
            "$ComputerNameStripped-UpsNet"
            "$ComputerNameStripped-sw"
            "$ComputerNameStripped-sw2"
            "$ComputerNameStripped-pbxUps"
            "$ComputerNameStripped-pbx"
            "$ComputerNameStripped-ups"
            "$ComputerNameStripped-vm-ilo"
            "$ComputerNameStripped-vm"
            "$ComputerNameStripped-db"
            "$ComputerNameStripped-ns"
            "$ComputerNameStripped-api"
            "$ComputerNameStripped-fuse"
            "$ComputerNameStripped-util"
        )
    }
    elseif ($Cluster) {
        $SiteHostnames = @(
            "$ComputerNameStripped-gw"
            "$ComputerNameStripped-UpsNet1"
            "$ComputerNameStripped-sw"
            "$ComputerNameStripped-UpsNet2"
            "$ComputerNameStripped-sw2"
    	    "$ComputerNameStripped-pbxUps"
            "$ComputerNameStripped-pbx"
            "$ComputerNameStripped-ups01"
            "$ComputerNameStripped-ups02"
            "$ComputerNameStripped-pdu01"
            "$ComputerNameStripped-pdu02"
            "$ComputerNameStripped-vm-ilo"
            "$ComputerNameStripped-vm"
        )
        $SiteHostnames = $SiteHostnames + $(if ($ClusterSite -notlike '*vSAN*') { @("$ComputerNameStripped-sanA", "$ComputerNameStripped-sanB") } ) +
        $Cluster.replace('.vcaantech.com', '-ilo') + $Cluster.replace('.vcaantech.com', '')
    }

    <# Revisit. Error checking text to see how the variable change as we override $Cluster for vCenter Cluster sites that moved to the cloud and no longer have ESXi hosts.  We still want those hospitals to make it to the multi-server site logic.
        Write-Host "`nIn function #2"
        Write-Host "ComputerName: $ComputerName"
        Write-Host "Cluster: $Cluster"
        Write-Host "ClusterSite: $ClusterSite"
        Write-Host "NetServices: $NetServices`n"
    #>

    if (Get-Module -Name ActiveDirectory) {
        Clear-Variable -Name ADComputers -ErrorAction Ignore
        $ADComputers = Get-ADComputer -Filter "Name -like '$ComputerNameStripped-*' -and Name -notlike '$ComputerNameStripped-*CNF:*' -and OperatingSystem -like '*Server*'" -Property IPv4Address |
            Where-Object IPv4Address -ne $null | Select-Object -ExpandProperty Name
        $SiteHostnames = ($SiteHostnames + $ADComputers +
            @(
                "$ComputerNameStripped-api"
                "$ComputerNameStripped-fuse"
                "$ComputerNameStripped-util"
            )
        ).ToLower() | Select-Object -Unique
    }
    else {
        Write-Warning 'ActiveDirectory module not found.'
        Write-Warning 'For enhanced hostname retrieval please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
    }

    Write-Output $SiteHostnames
}



function Get-IloStatus {
    # #27
    param(
        [string[]]$NetServices,
        [string[]]$Cluster,
        [string]$ComputerName,
        [pscredential]$Credential
    )
    if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
        if ((-not $Cluster) -and $NetServices) {
            $ServerIlo = "$ComputerName-ilo"
        }
        else {
            $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
        }

        if (-not $Credential) { $Credential = Get-Credential -Message "iLO Credentials:" }

        if ($Credential) {
            try {
                $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $Credential -DisableCertificateAuthentication -ErrorAction Stop

                $IloConnection | ForEach-Object {
                    Clear-Variable -Name ServerInfo -ErrorAction Ignore
                    $IloConnection_Item = $PSItem
                    $ServerInfo = Find-HPEiLO -Range $IloConnection_Item.IP -Full
                    $ServerSN = $ServerInfo.HostSystemInformation.SerialNumber
                    $IloObj = Get-HPEiLOHealthSummary -Connection $IloConnection_Item
                    $IloObj2 = Get-HPEiLOFan -Connection $IloConnection_Item
                    $IloObj3 = Get-HPEiLOPowerSupply -Connection $IloConnection_Item
                    $IloObj4 = Get-HPEiLOMemoryInfo -Connection $IloConnection_Item
                    $IloObj5 = Get-HPEiLOPowerRegulatorSetting -Connection $IloConnection_Item
                    $IloObj6 = Get-HPEBIOSProcessorPower -Connection $IloConnection_Item

                    $IloObjPowerProfile = switch -Wildcard ($IloConnection_Item.TargetInfo.ProductName) {
                        "*Gen8" { 'N/A' }
                        "*Gen9" { (Get-HPEBIOSPowerProfile -Connection $IloConnection_Item).PowerProfile }
                        "*Gen10" { (Get-HPEBIOSWorkloadProfile -Connection $IloConnection_Item).WorkloadProfile }
                    }

                    (Write-Output $IloObj | Select-Object -Property Hostname, IP, @{n = 'ProductName'; e = { $IloConnection_Item.TargetInfo.ProductName } },
                        @{n = 'SerialNumber'; e = { $ServerSN } }, BatteryStatus, StorageStatus | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n\r\n', ''

                    (Write-Output $IloObj | Select-Object -Property BIOSHardwareStatus, FanStatus, FanRedundancy, MemoryStatus, NetworkStatus, PowerSuppliesStatus,
                        PowerSuppliesRedundancy, ProcessorStatus, TemperatureStatus, @{n='PowerRegulatorMode';e={$IloObj5.Mode}}, @{n='EnergyPerformanceBias';e={$IloObj6.EnergyPerformanceBias}},
                        @{n='PowerProfile';e={$IloObjPowerProfile}} | Format-List | Out-String) -replace '\r\n\r\n\r\n\r\n', ''
                    (Write-Output $IloObj2.Fans | Select-Object -Property Name, Location, SpeedPercentage, State | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n\r\n', ''

                    # psu
                    if ($IloConnection_Item.TargetInfo.iLOGeneration -eq 'iLO5') {
                        (Write-Output $IloObj3.PowerSupplies | Select-Object -Property BayNumber, PowerSupplyStatus | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''
                    }
                    elseif ($IloConnection_Item.TargetInfo.iLOGeneration -eq 'iLO4') {
                        (Write-Output $IloObj3.PowerSupplies | Select-Object -Property Label, Status | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''
                    }

                    # memory
                    (Write-Output $IloObj4.MemoryDetailsSummary | Select-Object -Property * -ExcludeProperty NumberofSlots | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''

                    # nics
                    ($ServerInfo.HostSystemInformation.NICS | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n', ''

                    # Draw horizontal line for clusters but exclude drawing line on last host.
                    if ($IloConnection.Count -gt 1 -and $IloConnection.IndexOf($IloConnection_Item) -lt ($IloConnection.Count - 1)) { Write-Host "$('-'*85)`r" -ForegroundColor Cyan }
                }
                if ($Credential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                    New-StoredCredential -Credentials $Credential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                }
            }
            catch {
                Write-Warning $_.Exception.Message
                #Clear-Variable -Name IloCredential -Scope Global -ErrorAction Ignore
                Write-Host ''
            }
        }
    }
    else {
        Write-Warning "HPEiLOCmdlets module not found."
        Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
    }
}



function Get-IloBIOSPower {
    # #27p
    param(
        [string[]]$NetServices,
        [string[]]$Cluster,
        [string]$ComputerName,
        [pscredential]$Credential
    )
    if ((Get-Module -Name HPEiLOCmdlets) -or (Get-Module -Name HPEiLOCmdlets -ListAvailable)) {
        if ((-not $Cluster) -and $NetServices) {
            $ServerIlo = "$ComputerName-ilo"
        }
        else {
            $ServerIlo = $Cluster | ForEach-Object { "$($PSItem -replace '.vcaantech.com','')-ilo" }
        }

        if (-not $Credential) { $Credential = Get-Credential -Message "iLO Credentials:" }

        if ($Credential) {
            try {
                $IloConnection = Connect-HPEiLO -IP $ServerIlo -Credential $Credential -DisableCertificateAuthentication -ErrorAction Stop

                $IloConnection | ForEach-Object {
                    Clear-Variable -Name ServerInfo -ErrorAction Ignore
                    $IloConnection_Item = $PSItem
                    $ServerInfo = Find-HPEiLO -Range $IloConnection_Item.IP -Full
                    $ServerSN = $ServerInfo.HostSystemInformation.SerialNumber
                    $IloObj = Get-HPEiLOHealthSummary -Connection $IloConnection_Item
                    $IloObj5 = Get-HPEiLOPowerRegulatorSetting -Connection $IloConnection_Item
                    $IloObj6 = Get-HPEBIOSProcessorPower -Connection $IloConnection_Item

                    $IloObjPowerProfile = switch -Wildcard ($IloConnection_Item.TargetInfo.ProductName) {
                        "*Gen8" { 'N/A' }
                        "*Gen9" { (Get-HPEBIOSPowerProfile -Connection $IloConnection_Item).PowerProfile }
                        "*Gen10" { (Get-HPEBIOSWorkloadProfile -Connection $IloConnection_Item).WorkloadProfile }
                    }

                    (Write-Output $IloObj | Select-Object -Property Hostname, IP, @{n = 'ProductName'; e = { $IloConnection_Item.TargetInfo.ProductName } },
                        @{n = 'SerialNumber'; e = { $ServerSN } }, BatteryStatus, StorageStatus | Format-Table -AutoSize | Out-String) -replace '\r\n\r\n\r\n', ''

                    (Write-Output $IloObj | Select-Object -Property BIOSHardwareStatus, FanStatus, FanRedundancy, MemoryStatus, NetworkStatus, PowerSuppliesStatus,
                        PowerSuppliesRedundancy, ProcessorStatus, TemperatureStatus, @{n='PowerRegulatorMode';e={$IloObj5.Mode}}, @{n='PowerProfile';e={$IloObjPowerProfile}} |
                            Format-List | Out-String) -replace '\r\n\r\n\r\n\r\n', ''

                    if ($IloConnection_Item.TargetInfo.ProductName -notlike "*Gen8") {
                    (Write-Output $IloObj6 | Select-Object -Property Status, CollaborativePowerControl, DynamicPowerSavingsModeResponse, EnergyPerformanceBias,
                        IntelDMILinkFrequency, MinimumProcessorIdlePowerCoreState, MinimumProcessorIdlePowerPackageState | Format-List | Out-String) -replace '\r\n\r\n\r\n', ''
                    }
                    else {
                        Write-Host "`r`nGen8 power settings query not supported.`r`n" -ForegroundColor Cyan
                    }
                    # Draw horizontal line for clusters but exclude drawing line on last host.
                    if ($IloConnection.Count -gt 1 -and $IloConnection.IndexOf($IloConnection_Item) -lt ($IloConnection.Count - 1)) { Write-Host "$('-'*85)`r" -ForegroundColor Cyan }
                }
                if ($Credential -and (-not (Get-StoredCredential -Target vcahospilo))) {
                    New-StoredCredential -Credentials $Credential -Target vcahospilo -Type Generic -Persist LocalMachine | Out-Null
                }
            }
            catch {
                Write-Warning $_.Exception.Message
                Clear-Variable -Name IloCredential -Scope Global -ErrorAction Ignore
                Write-Host ''
            }
        }
    }
    else {
        Write-Warning "HPEiLOCmdlets module not found."
        Write-Warning "Please install by launching an elevated powershell session and entering:`nInstall-Module -Name HPEiLOCmdlets"
    }
}



function Get-GuestResource {
    # #11
    [CmdletBinding()]
    param(
        [parameter(Position = 0)]
        [string[]]$ComputerName,
        [pscredential]$Credential,
        [CimSession[]]$CimSession
    )
    # Proceed if server was selected
    Write-Host ''
    if ($ComputerName -or $CimSession) {
        if ($ComputerName -match '-dc' -and (-not $Credential)) { $Credential = Get-ADCreds }

        if (-not $CimSession) {
            foreach ($ComputerName_Item in $ComputerName) {
                try {
                    $CimSession = $CimSession + (New-CimSession -ComputerName $ComputerName_Item -Credential $Credential -ErrorAction Stop)
                }
                catch {
                    Write-Warning "[$ComputerName_Item] $($PSItem.Exception.Message)"
                    ''
                }
            }
        }

        if ($CimSession) {
            Get-MemoryUsage -CimSession $CimSession | Out-TableString -NoNewLine
            Get-DiskUsage -CimSession $CimSession | Out-TableString -NoNewLine

            # Retrieve cpu info and calculate load
            $CimSession | ForEach-Object {
                Clear-Variable -Name CpuLoad -ErrorAction Ignore
                $CimSession_Item = $PSItem
                try {
                    $CpuLoad = Get-CimInstance -CimSession $CimSession_Item -ClassName Win32_Processor -Property LoadPercentage -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average
                    Get-CimInstance -CimSession $CimSession_Item -ClassName Win32_Processor -Property Name, NumberOfCores, NumberOfLogicalProcessors -ErrorAction Stop |
                        Group-Object -Property NumberOfCores |
                        Select-Object -Property @{n = 'ComputerName'; e = { $CimSession_Item.ComputerName } },
                            @{n = 'CpuSockets'; e = { $_.Count } },
                            @{n = 'CoresPerSocket'; e = { [int]($_.Name) } },
                            @{n = 'TotalCPUs'; e = { $_.Count * $_.Name } },
                            @{n = 'CpuLoad(%)'; e = { [decimal]('{0:N2}' -f $CpuLoad.Average) } }
                }
                catch {
                    Write-Warning "[$($CimSession_Item.ComputerName)] CPU Core Check: $($PSItem.Exception.Message)"
                }
            } | Out-TableString -NoNewLine
            if ($CimSession.ComputerName -match '-util') {
                Get-CimInstance -CimSession $($CimSession | Where-Object { $_.ComputerName -like '*-util*'}) -ClassName Win32_ComputerSystem |
                    Select-Object -Property PSComputerName, Model, SystemSKUNumber, SystemType | Out-TableString -NoNewLine
                Get-CimInstance -CimSession $($CimSession | Where-Object { $_.ComputerName -like '*-util*'}) -ClassName Win32_BIOS |
                    Select-Object -Property PSComputerName, SerialNumber, SMBIOSBIOSVersion, ReleaseDate | Out-TableString -NoNewLine
            }
        }
        if ($CimSession) { Remove-CimSession -CimSession $CimSession -ErrorAction Ignore }

        if ($Credential -and (-not (Get-StoredCredential -Target vcadomaincreds))) {
            New-StoredCredential -Credentials $Credential -Target vcadomaincreds -Type Generic -Persist LocalMachine | Out-Null
        }
    }
}



#test new method
function Get-WWAPIHealth {
    # #37
    param(
    [pscredential]$Credential,
    [pscustomobject[]]$HospitalMaster
    )
    # WW API Health Report
    Clear-Variable -Name WWAPICheck, WWAPISelection, ADQueryResult -ErrorAction Ignore
    try {
    	$WWAPICheck = & "$PSScriptRoot\Private\bin\curl.exe" 'https://api.vcaantech.com/api/v1/health/check' | ConvertFrom-Json | Select-Object -ExpandProperty subHealthCheckResults |
            Where-Object Name -eq Hospitals | Select-Object -ExpandProperty subHealthCheckResults
    }
    catch {
        Write-Warning $_.Exception.Message
    }

    if ($WWAPICheck) {
        $WWAPICheck[1..($WWAPICheck.count -1)] | Start-RSJob -Name 'WWAPICheckJob' -FunctionsToImport 'Convert-VcaAU' -VariablesToImport HospitalMaster -Throttle 4 -ScriptBlock {
            $WWAPICheck_Item = $_
            $ErrorActionPreference = 'SilentlyContinue'
            $WWAPIServer = if ($WWAPICheck_Item.Name -like "*-api.*") { [System.Net.Dns]::GetHostEntry([regex]::Match($WWAPICheck_Item.Name,'(?i)h\d+-api').Value).HostName }
            $ErrorActionPreference = 'Continue'

            $WWAPIServerAU = Convert-VcaAU -AU $WWAPIServer -Strip -ErrorAction Ignore
            $SiteAUData = $HospitalMaster.Where({$PSItem.'Hospital Number' -eq $WWAPIServerAU})

            [pscustomobject]@{
                Server              = $WWAPIServer
                Status              = $WWAPICheck_Item.Status
                Name                = $WWAPICheck_Item.Name
                Message             = $WWAPICheck_Item.Message
                ElapsedMilliseconds = $WWAPICheck_Item.ElapsedMilliseconds
                TimeZone            = $SiteAUData.'Time Zone'
                City                = $SiteAUData.'City'
                State               = $SiteAUData.'St'
                HM                  = $SiteAUData.'Hospital Manager'
                Phone               = $SiteAUData.'Phone'
                SystemType          = $SiteAUData.'System Type'
                AU                  = "AU$WWAPIServerAU"
            }
        } | Out-Null
        Get-RSJob -Name 'WWAPICheckJob' | Wait-RSJob -ShowProgress -Timeout 300 | Receive-RSJob |
            Sort-Object -Property Status -Descending | Out-GridView -Title "#37 WOOFware API Health Report - Select site(s) to generate ServiceNow Case - $((Get-Date).ToString("yyyy-MM-dd HH:mm")) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Multiple -OutVariable WWAPISelection | Out-Null
        Get-RSJob -Name 'WWAPICheckJob' | Remove-RSJob

        # Generate ServiceNow Incident
        if ($WWAPISelection) {
            $UserEmail = (Get-ADUser -Identity $env:USERNAME -Properties Mail).Mail

            $MsgBoxInput = [System.Windows.MessageBox]::Show('Are you sure?', 'Generate ServiceNow Incident', 'YesNo', 'Question')
            switch ($MsgBoxInput) {
                'Yes' {
                    $WWAPISelection | ForEach-Object {
                        $WWAPISNOWAU = Convert-VcaAU -AU $PSItem.Server -Strip -ErrorAction SilentlyContinue
                        $WWAPISNOWTZ = $(($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq $WWAPISNOWAU } )).'Time Zone')
                        $WWAPIShortDescription = "AU$(Convert-VcaAU -AU $PSItem.Server -Strip) - $($PSItem.Server) - WOOFware API Warning - TZ: $WWAPISNOWTZ"

                        Clear-Variable -Name WWAPIResult -ErrorAction Ignore
                        $NewServiceNowParams = @{
                            ImpactedUser     = $(($HospitalMaster.Where( { $PSItem.'Hospital Number' -eq $WWAPISNOWAU } )).'Hospital Manager Email') #String email address of SNOW user:Default Guest
                            ReportedBy       = $UserEmail     #String email address of SNOW user:Default Guest

                            Category         = "Network"      #String accepted values:"hardware";"inquiry";"network";"software";"database";"security":Default "inquiry"
                            CIName           = "Router - ISP" #String:Name of CI:Default "Hardware - Other"
                            Impact           = "2"            #Numeric:Range 1 to 3:Default 3
                            Urgency          = "2"            #Numeric:Range 1 to 3:Default 3

                            ContactType      = "Monitoring"     #String: accepted values: "messenger";"email";"phone";"self-service";"monitoring";"voice mail";"walk-in/direct-contact":DEFAULT "monitoring"
                            AssignedGroup    = "VCA Operations" #String:Name of SNOW Group:Default "Support Alerts"
                            AssignedTo       = ""             #Nameof of a member of SNOW group:Default NULL if AssignedTo user is not a member of the AssignedGroup

                            ShortDescription = "$WWAPIShortDescription" #String: Default value "Short Description is MISSING"
                            Description      = "WOOFware API Health Report created by Ops Portal Tool v.$($Version): $($PSItem | Format-List | Out-String)" #String: Default value "Description is MISSING"
                        }
                        if ($Credential) { $NewServiceNowParams.Credential = $Credential }
                        $WWAPIResult = New-ServiceNowIncident @NewServiceNowParams

                        if ($WWAPIResult) {
                            Write-Host "`r`nCase Generated:`r $($WWAPIResult.incident) - $WWAPIShortDescription`r`n" -ForegroundColor Yellow
                            $PSItem | Out-TableString
                            Write-Host "`r`n$($WWAPIResult.url)" -ForegroundColor Cyan
                            Start-Process $WWAPIResult.url
                        }
                        Write-Host ''
                    }
                }
                'No' {
                }
            } #MsgBox Switch Case
        } #generate snow case
    }
}



function Invoke-SnowGui {
    # #14t
    param (
        [string]$ComputerName,
        [pscustomobject[]]$HospitalMaster,
        [string]$ImpactedUser,
        [string[]]$ImpactedUserList,
        [pscredential]$SNOWAPICredential
    )
    if ($HospitalMaster) {
        $HospitalInfo = $HospitalMaster.Where( {
                $PSItem.'Hospital Number' -eq "$(Convert-VcaAU -AU $ComputerName -Strip)"
            } )
        if ($HospitalInfo) {
            Write-Host 'Location:'
            Write-Host "$($HospitalInfo.'Operating Name') #$($HospitalInfo.'Hospital Number')"
            Write-Host "$($HospitalInfo.Address)"
            Write-Host "$($HospitalInfo.City), $($HospitalInfo.St) $($HospitalInfo.Zip)"
            Write-Host ''
            Write-Host 'VCA Site Contact:'
            Write-Host "$($HospitalInfo.'Hospital Manager'), $($HospitalInfo.'Hospital Manager Email')"
            Write-Host "$($HospitalInfo.Phone)"
            Write-Host ''
            Write-Host 'Misc. info:'
            Write-Host 'Time Zone              :'"$($HospitalInfo.'Time zone')"
            Write-Host 'URL                    :'"$($HospitalInfo.GPURL)"
            Write-Host 'Back Line              :'"$($HospitalInfo.'Back Line')"
            Write-Host 'System Conversion Date :'"$($HospitalInfo.'System Conversion Date')"
            Write-Host 'System Type            :'"$($HospitalInfo.'System Type')"
        }
        Write-Host ''
    }
    if (-not $SNOWAPICredential) { $SNOWAPICredential = Get-StoredCredential -Target vcasnowapi }
    if (-not $ImpactedUser) { $ImpactedUser = "$($HospitalInfo.'Hospital Manager Email')" }

    $SiteHostnames = Get-VcaSiteHostname -ComputerName $ComputerName -Cluster $Cluster -ClusterSite $ClusterSite -NetServices $NetServices
    $SiteHostnames = $SiteHostnames -notmatch '-vmincom|-vmdecom'

    $ServiceNowParams = @{
        ImpactedUser     = $ImpactedUser
        ImpactedUserList = $ImpactedUserList
        CIName           = $ComputerName
        ComputerName     = $SiteHostnames
        ShortDescription = "AU$(Convert-VcaAU -AU $ComputerName -Strip) - $ComputerName - No Description"
    }
    if ($SNOWAPICredential) { $ServiceNowParams.Credential = $SNOWAPICredential }

    New-ServiceNowGUI @ServiceNowParams
}



function Get-VcaHospitalMaster {
    [CmdletBinding()]
    param(
        $SharePointUrl = 'https://vca365.sharepoint.com/sites/WOOFconnect/regions'
    )
    try {
        Connect-PnPOnline -Url $SharePointUrl -UseWebLogin -ErrorAction Stop -WarningAction Ignore
        Get-PnPFile -Url '/Documents/HOSPITALMASTER.xlsx' -Path "$PSScriptRoot\private\csv" -Filename 'HOSPITALMASTER_new.xlsx' -AsFile -Force -ErrorAction Stop
    }
    catch {
        throw $_.Exception.Message
    }
}



function Update-HospitalMaster {
    # #14u
    param(
        [pscredential]$EmailCredential
    )
    # Check for Hospital Master update
    if (-not $EmailCredential) { $EmailCredential = Get-StoredCredential -Target vcaemailcreds }

    $CsvPath = "$PSScriptRoot\private\csv"
    $HospitalMasterXlsx = "$CsvPath\HOSPITALMASTER.xlsx"
    $HospitalMasterXlsxNew = "$CsvPath\HOSPITALMASTER_new.xlsx"

    if (-not (Test-Path -Path $CsvPath)) { New-Item -ItemType Directory -Path $CsvPath | Out-Null }

    # Download CSV
    try {
        if (Test-Path -Path $HospitalMasterXlsx) {
            try {
                Get-VcaHospitalMaster -ErrorAction Stop

                # Get file hash
                $CurrentHash = Get-FileHash -Path $HospitalMasterXlsx -Algorithm SHA256
                $NewHash = Get-FileHash -Path $HospitalMasterXlsxNew -Algorithm SHA256

                # Check if downloaded CSV is newer (different)
                if ($CurrentHash.Hash -ne $NewHash.Hash) {
                    Write-Host "New version of hospital master found... updating`n" -ForegroundColor Cyan
                    if (Test-Path -Path $HospitalMasterXlsxNew) {
                        # Rename and move new file
                        Move-Item -Path $HospitalMasterXlsxNew -Destination $HospitalMasterXlsx -Force

                        if ((Read-Choice -Title "Save copy to network share?" -DefaultChoice 1) -eq 1) {
                            Copy-Item -Path "$HospitalMasterXlsx" -Destination "\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\HOSPITALMASTER.xlsx" -Confirm:$false -Force -Verbose
                        }
                    }
                }
                else {
                    if (Test-Path -Path $HospitalMasterXlsxNew) {
                        $HospitalFileDate = '{0:M/dd/yyyy h:mm tt}' -f (Get-Item -Path $HospitalMasterXlsx | Select-Object -ExpandProperty LastWriteTime)
                        Write-Host "Hospital master XLSX is already up-to-date (Last Write Time: $HospitalFileDate)`n" -ForegroundColor Cyan

                        # Remove duplicate file
                        Remove-Item -Path $HospitalMasterXlsxNew
                    }
                }
            }
            catch {
                Write-Warning $_.Exception.Message
            }
        }
        else {
            Write-Host "Hospital master xlsx not found... downloading file`n" -ForegroundColor Cyan
            Get-VcaHospitalMaster -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "[Download failed] $($PSItem.Exception.Message)"
        if (Test-Path -Path $HospitalMasterXlsx) {
            $HospitalFileDate = '{0:M/dd/yyyy h:mm tt}' -f (Get-Item -Path $HospitalMasterXlsx | Select-Object -ExpandProperty LastWriteTime)
            Write-Host "Using local hospital master XLSX (Last Write Time: $HospitalFileDate)`n" -ForegroundColor Cyan
        }
    }
}



function Invoke-WWAPIVerification {
    # #37c
    param(
        [pscustomobject[]]$HospitalMaster
    )
    $WWAPICheck = Invoke-RestMethod -Uri 'https://api.vcaantech.com/api/v1/health/check' -UseBasicParsing | Select-Object -ExpandProperty subHealthCheckResults |
        Where-Object Name -eq Hospitals | Select-Object -ExpandProperty subHealthCheckResults

    if (-not $WWAPICheck) { Write-Warning "Could not retrieve WOOFware API Report." ; break }
    # Retrieve AU number from API report
    $WWAPIResults =
    foreach ($WWAPICheck_Item in $WWAPICheck) {
        $WWAPIServerAU = $(Convert-VcaAU -AU ([regex]::Match($WWAPICheck_Item.Name,'(?i)h\d+-api').Value) -Strip -ErrorAction SilentlyContinue)
        [pscustomobject]@{
            AU = $WWAPIServerAU
        }
    }
    # Filter hospital master for woofware converted sites
    $HospitalMasterWW = $HospitalMaster.Where({ $PSItem.'System Type' -eq 'WOOFware' -and $PSItem.'System Conversion Date' })

    # Find missing AUs in API Report
    Write-Host "WOOFware API Verification $(Get-Date -Format g)`r"
    Write-Host "Hospital Master sites missing from API Report:" -ForegroundColor Cyan
    $MismatchResults = $HospitalMasterWW.Where({ $PSItem.'Hospital Number' -notin $WWAPIResults.AU })
    if (-not $MismatchResults) { Write-Host "`r`n* No missing entries found`r`n" -ForegroundColor Yellow }
    else { ($MismatchResults | Out-String) -replace '\r\n\r\n', '' }

    Write-Host "Hospital Master WW sites: $($HospitalMasterWW.count)" -ForegroundColor Cyan
    Write-Host "WOOFware API Report Count: $($WWAPIResults.count)`r`n" -ForegroundColor Cyan

    Write-Host "$('-'*43)"

    # Find missing AUs in Hospital Master
    Write-Host " API entries missing from Hospital Master:" -ForegroundColor Cyan
    $MismatchResults2 = $WWAPIResults.Where({ $PSItem.AU -notin $HospitalMasterWW.'Hospital Number' })
    if ($MismatchResults2) { ($MismatchResults2 | Select-Object -Property @{n='AU';e={$_.AU -as [int]}} | Sort-Object -Property AU | Out-String) -replace '\r\n\r\n', '' }
    else { Write-Host "`r`n* No missing entries found`r`n" -ForegroundColor Yellow }

    Write-Host "Missing API Entries: $($WWAPIResults.count - $HospitalMasterWW.count)`r`n" -ForegroundColor Cyan
}



function Select-VcaSite {
    param(
        [string[]]$AU,
        [string]$Title,
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Single', 'Multiple', 'None')]
        [string]$OutputMode = 'Multiple'
    )

    $SiteAU = Convert-VcaAu -AU $AU -Suffix ''

    # Gather Site AD Objects and show user selection screen
    Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties IPv4Address, OperatingSystem |
        Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'Status'; e = { $PSItem.Name | Get-PingStatus } } | Sort-Object -Property Name |
        #Select-Object -Property Name, IPv4Address, OperatingSystem, @{n = 'PingStatus'; e = { (Test-ConnectionAsync -ComputerName $PSItem.Name -Full).Result } } | Sort-Object -Property Name |
            Out-GridView -Title $Title -OutputMode $OutputMode
}



function Get-VcaADComputers {
    # #20
    param(
        [string[]]$AU
    )
    # active directory computers
    if (Get-Module -Name ActiveDirectory) {
        $SiteAU = Convert-VcaAu -AU $AU -Suffix ''

        if ($SiteAU) {
            $ADComputers = Get-ADComputer -Filter "Name -like '$SiteAU-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -or Name -like '$SiteAU-Util*'" -Properties OperatingSystem, IPv4Address, CanonicalName
        }
        else {
            $ADComputers = Get-ADComputer -Filter "Name -like '$AU' -and Name -like 'hmtprod-*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*'" -Properties OperatingSystem, IPv4Address, CanonicalName
        }
        $ADComputers | Select-Object -Property Name, IPv4Address, OperatingSystem, CanonicalName,
                @{n='Subnet';e={
                    if ($_.IPv4Address -like "10.242.*") {
                        if (Get-IPAddressSubnetMember -IPAddress $_.IPv4Address -StartRange '10.242.72.0' -EndRange '10.242.103.255') {
                            'VMC West'
                        }
                        elseif (Get-IPAddressSubnetMember -IPAddress $_.IPv4Address -StartRange '10.242.136.0' -EndRange '10.242.168.255') {
                            'VMC East'
                        }
                        else {
                            'VMC'
                        }
                    }
                    elseif ($_.IPv4Address -like "10.225.*") { 'Mesa' }
                }} | Sort-Object -Property CanonicalName
    }
    else {
        Write-Warning 'ActiveDirectory module not found.'
        Write-Warning 'Please install RSAT https://www.microsoft.com/en-us/download/details.aspx?id=45520'
    }
}



function Get-IPAddressSubnetMember {
    param(
        [version]$IPAddress,
        [version]$StartRange,
        [version]$EndRange
    )
    $IPAddress -gt $StartRange -and
    $IPAddress -lt $EndRange
}



function Update-ClusterCsv {
    # #14u
    [cmdletbinding()]
    param(
        [string[]]$Server = @('phvccl01.vcaantech.com','phvcahvcp01.vcaantech.com'),
        [pscustomobject[]]$Clusters,
        [parameter(Mandatory)]
        [pscredential]$Credential
    )
    try {
        Clear-Variable -Name ClusterSitesNew, ClusterCompare -ErrorAction Ignore
        Write-Verbose "Connecting to vCenter..."
        $VIServer = Connect-VIServer -Server $Server -Credential $Credential -ErrorAction Stop

        Write-Verbose "Retrieving clusters list..."
        $ClusterException = @(
            #'H0938'
            'H0987'
            'H1143'
            'H1144'
            'H1327'
            'H4022'
        )
        $RegExFilter = ($ClusterException | foreach-object {[regex]::escape($_)}) -join '|'
        $ClusterSitesNew = Get-Cluster -Name h* -Server $VIServer | Where-Object {$_ -notmatch $RegExFilter} | Get-VMHost |
            Select-Object -Property @{n='AU';e={Convert-VcaAU $_.Parent.Name.Split(' ')[0] -Strip}}, Name, @{n='Cluster';e={$_.Parent}} |
            Sort-Object -Property Name
        Disconnect-VIServer -Server $Server -Force -Confirm:$false

        if ($ClusterSitesNew) {
            Write-Verbose "Comparing lists..."
            $ClusterCompare = Compare-Object -ReferenceObject $Clusters -DifferenceObject $ClusterSitesNew -Property Name
            if ($ClusterCompare) {
                if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
                    Write-Host "Found below changes:" -ForegroundColor Cyan
                    Write-Host ' => - New' -ForegroundColor Cyan
                    Write-Host ' <= - Removed' -ForegroundColor Cyan
                    $ClusterCompare | Out-String
                    Write-Host "Saving new list." -ForegroundColor Cyan
                }
                $script:ClusterSites = $ClusterSitesNew
                $ClusterSitesNew | Export-Csv -Path "$PSScriptRoot\Private\csv\ClusterSites.csv" -NoTypeInformation -Force

                if ((Read-Choice -Title "Save copy to network share?" -DefaultChoice 1) -eq 1) {
                    Copy-Item -Path "$PSScriptRoot\Private\csv\ClusterSites.csv" -Destination "\\vcaantech.com\folders\data2\corp\Information Technology\Operations\Projects\Scripting\VCA Ops Portal\Private\csv\ClusterSites.csv" -Confirm:$false -Force -Verbose
                }
            }
            else {
                Write-Verbose "No changes in clusters list found."
            }
        }
    }
    catch {
        Write-Warning "Error [Connecting to vCenter/ESXi]: $_`n"
    }
}



function Invoke-CachedPwRefresh {
    # #99r
    $script:EsxiCredential = Get-StoredCredential -Target vcahospesxi
    $script:IloCredential = Get-StoredCredential -Target vcahospilo
    $script:ADCredential = Get-StoredCredential -Target vcadomaincreds
    $script:SNOWAPICredential = Get-StoredCredential -Target vcasnowapi
    $script:EmailCredential = Get-StoredCredential -Target vcaemailcreds
    Write-Host "`r`nCached passwords refreshed`r`n" -ForegroundColor Cyan
}



function Invoke-CachedListsRefresh {
    # Load hospital master to memory
    $script:HospitalMaster = Import-Excel -Path "$PSScriptRoot\Private\csv\HOSPITALMASTER.xlsx" -WorksheetName Misc
    # Load clusters list to memory
    $script:ClusterSites = Import-Csv -Path "$PSScriptRoot\Private\csv\ClusterSites.csv"
    # Load Ops Full Menu Csv
    $script:PortalMenuCsv = Import-Csv -Path "$PSScriptRoot\Private\lib\Menu.csv"
    # Circuit Site Selection
    $script:HospitalCircuits = Import-Csv -Path "$PSScriptRoot\Private\csv\All-Hospital-Circuits.csv"

    Write-Host "`r`nCached Lists Refreshed`r`n" -ForegroundColor Cyan
}



function Get-TSTime {
    # #11time
    param(
        [string[]]$ComputerName
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $DateTime = Get-Date
        [pscustomobject]@{
            'DateTime (Local)' = "{0:M/dd/yyyy h:mm tt} $([Regex]::Replace([System.TimeZoneInfo]::Local.Id, '([A-Z])\w+\s*', '$1'))" -f $DateTime
            'DateTime (UTC)'   = $DateTime.ToUniversalTime().ToString("M/dd/yyyy h:mm tt")
        }
    } -ErrorAction SilentlyContinue | Select-Object -Property PSComputerName, 'DateTime (Local)', 'DateTime (UTC)' | Sort-Object -Property PSComputerName
}



function Get-VcaStdTSNames {
    param(
        [string]$ComputerName
    )
    $SiteAU = Convert-VcaAu -AU $ComputerName -Suffix ''
    Get-ADComputer -Filter "Name -like '$SiteAU-fs*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -and Enabled -eq '$true'"
    Get-ADComputer -Filter "Name -like '$SiteAU-db*' -and Name -notlike '*CNF:*' -and OperatingSystem -like '*Server*' -and Enabled -eq '$true'"
}



function Get-WWUserMemory {
    # #11v
    param(
        [string[]]$ComputerName
    )
    try {
        $WWMemoryResults = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-Process -Name VCA.Sparky.Shell -IncludeUserName | Select-Object -Property WS, UserName, ProcessName, StartTime, ID
        } -ErrorAction SilentlyContinue
        $WWMemoryResults | Sort-Object -Property PSComputerName, @{Expression = "WS"; Descending = $True} |
            Select-Object -Property @{n='MemoryMB';e={'{0:N2}' -f ($_.WS/1MB)}}, Username, @{n='MemoryGB';e={'{0:N2}' -f ($_.WS/1GB)}}, ProcessName, StartTime, ID, PSComputerName
    }
    catch {
        Write-Warning $_.Exception.Message
    }
}



function Invoke-UpdateVcaCircuitsCsv {
    # #15
    $HospitalCsvPath = '\\laitocsvp01\csv\All-Hospital-Circuits.csv'
    $CsvPath = "$PSScriptRoot\Private\csv"
    $HospitalCsv = "$CsvPath\All-Hospital-Circuits.csv"
    $HospitalCsvNew = "$CsvPath\All-Hospital-Circuits_new.csv"

    if (-not (Test-Path -Path $CsvPath)) { New-Item -ItemType Directory -Path $CsvPath | Out-Null }

    # Download CSV
    try {
        if (Test-Path -Path $HospitalCsv) {
            Copy-Item -Path $HospitalCsvPath -Destination $HospitalCsvNew -ErrorAction Stop

            # Get file hash
            Clear-Variable -Name CurrentHash, NewHash -ErrorAction Ignore
            $CurrentHash = Get-FileHash -Path $HospitalCsv -Algorithm SHA256
            $NewHash = Get-FileHash -Path $HospitalCsvNew -Algorithm SHA256

            # Check if downloaded CSV is newer (different)
            if ($CurrentHash.Hash -ne $NewHash.Hash) {
                Write-Host 'New version of hospital circuits found... updating' -ForegroundColor Cyan
                if (Test-Path -Path $HospitalCsvNew) {
                    # Rename and move new file
                    Move-Item -Path $HospitalCsvNew -Destination $HospitalCsv -Force
                }
            }
            else {
                if (Test-Path -Path $HospitalCsvNew) {
                    # Remove duplicate file
                    Remove-Item -Path $HospitalCsvNew
                }
            }
        }
        else {
            Write-Host 'Hospital circuits csv not found... copying file' -ForegroundColor Cyan
            Copy-Item -Path $HospitalCsvPath -Destination $HospitalCsv -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "[Copy failed] $($_.Exception.Message)"
        if (Test-Path -Path $HospitalCsv) {
            $HospitalFileDate = '{0:M-dd-yyyy}' -f (Get-Item -Path $HospitalCsv | Select-Object -ExpandProperty LastWriteTime)
            Write-Host "Reading from local hospital circuits csv ($HospitalFileDate)" -ForegroundColor Cyan
        }
    }
}



function Get-VcaHospitalCircuits {
    # #15
    # Read CSV
    $CsvPath = "$PSScriptRoot\Private\csv"
    $HospitalCsv = "$CsvPath\All-Hospital-Circuits.csv"
    if (Test-Path -Path $HospitalCsv) {
        $script:HospitalCircuits = Import-Csv -Path $HospitalCsv
        $SiteAU = Convert-VcaAu -AU $ComputerName -Strip

        $HospitalCircuitMatch = $HospitalCircuits.Where( { $PSItem.AU -like "*$SiteAU*" } )

<# The is the depreciated code that was commented out in v.20250910
        if (@($HospitalCircuitMatch).count -gt 1) {
            $HospitalCircuitRegExMatch = $HospitalCircuits.Where( { $PSItem.AU -match "^$SiteAU[A-Za-z]?$|^$SiteAU[A-Za-z]?? |AU$SiteAU ?" } )
            if (@($HospitalCircuitRegExMatch).count -ne @($HospitalCircuitMatch).count) {
                Write-Host "RegEx Filter: `"^$SiteAU[A-Za-z]?$|^$SiteAU[A-Za-z]?? |AU$SiteAU ?`"" -ForegroundColor Cyan
                "`r`n$((Get-Date | Out-String).Trim())`r`n"
                Write-Output "$(($HospitalCircuitRegExMatch | Format-List | Out-String).Trim())`r`n"
            }
        }
#>

        $HospitalCircuitRegExMatch = $HospitalCircuits.Where( { $PSItem.AU -match "^$SiteAU[A-Za-z]?$|^$SiteAU[A-Za-z]?? |AU$SiteAU ?" } )

        Write-Host "`nCircuits URL: http://laitocsvp01/cgi-bin/x.All-Hospital-Circuits.cgi"
        Write-Host ""
        Write-Host "RegEx Filter: `"^$SiteAU[A-Za-z]?$|^$SiteAU[A-Za-z]?? |AU$SiteAU ?`"" -ForegroundColor Cyan
        Write-Host ""

        Write-Output "$(($HospitalCircuitRegExMatch | Format-List | Out-String).Trim())`r`n"

    } 
    else {
        Write-Warning "Hospital Circuits CSV not found, please make sure you're connected to VCA network."
    }
}



function Out-TableString {
    param(
        [parameter(
            ValueFromPipeline,
            Position = 0)]
        $InputObject,
        $Property,
        [switch]$NoAutosize,
        [switch]$Wrap,
        [switch]$NoNewLine
    )
    begin {
        $FormatTableParams = @{
            Property = $Property
            Autosize = $true
        }
        if ($Wrap.IsPresent) { $FormatTableParams.Wrap = $true }
        if ($NoAutosize.IsPresent) { $FormatTableParams.Autosize = $false }

        [System.Collections.ArrayList]$InputObject_All = @()
    }
    process {
        foreach ($InputObject_Item in $InputObject) {
            $InputObject_All.Add($InputObject_Item) | Out-Null
        }
    }
    end {
        if (-not $NoNewLine.IsPresent -and $InputObject_All -ne '' -and $null -ne $InputObject_All) {
            "`r`n$(($InputObject_All | Format-Table @FormatTableParams | Out-String).Trim())`r`n"
        }
        elseif ($NoNewLine.IsPresent -and $InputObject_All -ne '' -and $null -ne $InputObject_All) {
            "$(($InputObject_All | Format-Table @FormatTableParams | Out-String).Trim())`r`n"
        }
    }
}



function Out-ListString {
    param(
        [parameter(
            ValueFromPipeline,
            Position = 0)]
        $InputObject,
        $Property,
        [switch]$NoNewLine
    )
    begin {
        [System.Collections.ArrayList]$InputObject_All = @()
    }
    process {
        foreach ($InputObject_Item in $InputObject) {
            $InputObject_All.Add($InputObject_Item) | Out-Null
        }
    }
    end {
        if (-not $NoNewLine.IsPresent -and $InputObject_All -ne '' -and $null -ne $InputObject_All) {
            "`r`n$(($InputObject_All | Format-List -Property $Property | Out-String).Trim())`r`n"
        }
        elseif ($NoNewLine.IsPresent -and $InputObject_All -ne '' -and $null -ne $InputObject_All) {
            "$(($InputObject_All | Format-List -Property $Property | Out-String).Trim())`r`n"
        }
    }
}




function Get-UpsSnmp {
    #31s
    [CmdletBinding()]
    param(
        [string[]]$UPS,
        [string]$Community = 'vcapublic',
        [string[]]$Oid,
        [string]$Version = 'Ver1',
        [switch]$SnmpFallback,
        $HPUpsSnmp
    )
    $UPS | ForEach-Object {
        try {
            if (-not ($PSItem | Get-PingStatus)) {
                Write-Warning "[$PSItem] Timed out"
                # go to next hostname in foreach loop
                return
            }
            $UPS_Item = $PSItem
            $UpsIp = (Resolve-DnsName -Name $PSItem -ErrorAction Stop).IP4Address
            $SnmpResults = Invoke-SnmpGet -IpAddress $UpsIp -Community $Community -Oid $HPUpsSnmp.SnmpOid -Version $Version -ErrorAction Stop

            $SnmpResultsObj = [pscustomobject]::new()
            $SnmpResults | ForEach-Object {
                $SnmpResults_Item = $_
                $OidDescription = ($HPUpsSnmp | Where-Object SnmpOid -eq $SnmpResults_Item.OID).Description
                $SnmpResultsObj | Add-Member -Name $OidDescription -MemberType NoteProperty -Value $_.Value

                switch ($OidDescription) {
                    'Status' {
                        switch ($SnmpResults_Item.Value) {
                            '1' { $SnmpResultsObj.Status = "Other"; break }
                            '2' { $SnmpResultsObj.Status = "None"; break }
                            '3' { $SnmpResultsObj.Status = "Normal"; break }
                            '4' { $SnmpResultsObj.Status = "Bypass"; break }
                            '5' { $SnmpResultsObj.Status = "Battery"; break }
                            '6' { $SnmpResultsObj.Status = "Booster"; break }
                            '7' { $SnmpResultsObj.Status = "Reducer"; break }
                            '8' { $SnmpResultsObj.Status = "Parallel Capacity"; break }
                            '9' { $SnmpResultsObj.Status = "Parallel Redundant"; break }
                            '10' { $SnmpResultsObj.Status = "High Efficiency Mode" }
                        }
                        break
                    }
                    'Battery ABM Status' {
                        switch ($SnmpResults_Item.Value) {
                            '1' { $SnmpResultsObj.'Battery ABM Status' = "Battery Charging"; break }
                            '2' { $SnmpResultsObj.'Battery ABM Status' = "Battery Discharging"; break }
                            '3' { $SnmpResultsObj.'Battery ABM Status' = "Battery Floating"; break }
                            '4' { $SnmpResultsObj.'Battery ABM Status' = "Battery Resting"; break }
                            '5' { $SnmpResultsObj.'Battery ABM Status' = "unknown" }
                        }
                        break
                    }
                    'Battery Time Remaining' { if ($SnmpResults_Item.Value -match '\d') { $SnmpResultsObj.$_ = "$('{0:N0}' -f ($SnmpResults_Item.Value / 60)) Minutes" } }
                }
            }
            $SnmpResultsObj
        }
        # DNS Resolve failure
        catch [System.ComponentModel.Win32Exception] {
            Write-Warning "[$UPS_Item] $($_.Exception.Message)"
        }
        catch {
            #Write-Warning "[$UPS_Item] $($_.Exception.Message)"
            if ($SnmpFallback.IsPresent) {
                Write-Host "`r`n[$UPS_Item] SNMPv1 failed, trying SNMPv2." -ForegroundColor Cyan
                Get-UpsSnmp -UPS $UPS_Item -Version Ver2 -HPUpsSnmp $HPUpsSnmp
            }
            else {
                Write-Host "`r`n[$UPS_Item] SNMPv2 failed, verify SNMP Settings." -ForegroundColor Cyan
            }
        }
    }
}



function Invoke-WinRMQuickCfg {
    param(
        [string[]]$ComputerName,
        [pscredential]$Credential
    )
    $CimSessionParams = @{
        ComputerName  = $ComputerName
        SessionOption = New-CimSessionOption -Protocol Dcom
    }
    if ($Credential) { $CimSessionParams.Credential = $Credential }
    $CimSession = New-CimSession @CimSessionParams

    $cimParam = @{
        CimSession = $CimSession
        ClassName  = 'Win32_Process'
        MethodName = 'Create'
        Arguments  = @{ CommandLine = 'cmd.exe /c winrm quickconfig -quiet' }
    }
    Invoke-CimMethod @cimParam
}




function Find-WindowsUpdateInitiator {
    # #70u
    param(
        [string]$ComputerName,
        [pscredential]$Credential
    )
    $windowsupdatelogs = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        (Get-WindowsUpdateLog -LogPath 'C:\temp\WindowsUpdate.log') | Out-Null
        Select-String -Path 'C:\temp\WindowsUpdate.log' -Pattern '(?i)S-\d-\d+-(\d+-){1,14}\d+'
    } -Credential $Credential

    $windowsupdatelogs | Select-Object -Property PSComputerName, @{n='Result';e={$_.Line}},
        @{n='ADUserName';e={(Get-ADUser -Filter "SID -like '$([regex]::Match($_.Line, '(?i)S-\d-\d+-(\d+-){1,14}\d+').Value)'").SamAccountName}},
        @{n='DateTime';e={Get-Date ([regex]::Match($_.Line, '(0|‎)\d{1,2}\/(0|‎)\d{1,2}\/(0|‎)\d{1,4}').Value)}}
}



function Invoke-DhcpPrompt { #24b
    [CmdletBinding()]
    param (
        $ComputerName,
        $DhcpServer,
        $DhcpScopes,
        $DhcpScopeId,
        $DhcpScopeName,
        [pscredential]$Credential
    )
    $SitePrefix = Convert-VcaAu -AU $ComputerName -Strip

    try {
        if (-not $DhcpScopes -and -not $DhcpScopeId) {
            Write-Host "`r`n > Invoke-Command -ComputerName $($DhcpServer) `{ Get-DhcpServerv4Scope | Where-Object `{ `$_.Name -like `"*AU$SitePrefix-*`" `} `}" -ForegroundColor Cyan
            $DhcpScopeResults = Invoke-Command -ComputerName $DhcpServer { Get-DhcpServerv4Scope | Where-Object { $_.Name -like "*AU$using:SitePrefix-*" } } -Credential $Credential -ErrorAction Stop |
                Select-Object -Property ScopeId, SubnetMask, Name, State, StartRange, EndRange, LeaseDuration
            $DhcpScopeResults | Out-GridView -Title "#24b Select DHCP scope to query - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable DhcpScopeSelection | Out-Null
        } # scope selection
        elseif ($DhcpScopes -and -not $DhcpScopeId) {
            $DhcpScopes | Out-GridView -Title "#24b Select DHCP scope to query - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")" -OutputMode Single -OutVariable DhcpScopeSelection | Out-Null
        } # refresh leases
        else {
            $DhcpScopeSelection = @{
                ScopeId   = $DhcpScopeId
                ScopeName = $DhcpScopeName
            }
        }

        # show dhcp leases
        if ($DhcpScopeSelection) {
            if (-not $DhcpScopeId) {
                Write-Host " > Get-DhcpServerv4Lease -ComputerName $($DhcpServer) -ScopeId $($DhcpScopeSelection.ScopeId) #Scope Name: $($DhcpScopeSelection.Name)`r`n" -ForegroundColor Cyan
            }
                $SelectObjectParams = @{
                Property = @(
                    'IPAddress'
                    'HostName'
                    'LeaseExpiryTime'
                    'ClientType'
                    'ClientId'
                    $(@{n='*Vendor';e={(Select-String -Path "$PSScriptRoot\private\csv\manuf.txt" -Pattern $_.ClientId.Substring(0,8).replace('-',':')).line.Split("`t")[2]}})
                    'DnsRegistration'
                    'DnsRR'
                    'Description'
                    'NapStatus'
                    'PolicyName'
                    )
            }
            $DhcpLeaseResults = Get-DhcpServerv4Lease -ComputerName $DhcpServer -ScopeId $DhcpScopeSelection.ScopeId -ErrorAction Stop |
                Select-Object @SelectObjectParams | Sort-Object -Property IPAddress

            $RefreshDhcpLeases = '' | Select-Object -Property @(
                'IPAddress'
                @{n='HostName';e={'---------- Refresh DHCP Leases ----------'}}
                'LeaseExpiryTime'
                'ClientType'
                'ClientId'
                '*Vendor'
                'DnsRegistration'
                'DnsRR'
                'Description'
                'NapStatus'
                'PolicyName'
                )
            $ReselectDhcpScopes = '' | Select-Object -Property @(
                'IPAddress'
                @{n='HostName';e={'---------- DHCP Scope Selection ----------'}}
                'LeaseExpiryTime'
                'ClientType'
                'ClientId'
                '*Vendor'
                'DnsRegistration'
                'DnsRR'
                'Description'
                'NapStatus'
                'PolicyName'
                )
            $DhcpLeaseResults = @($RefreshDhcpLeases) + @($ReselectDhcpScopes) + @($DhcpLeaseResults)
            $DhcpLeaseResultsSelection = $DhcpLeaseResults | Out-GridView -PassThru -Title "#24b AU$SitePrefix - DhcpServerv4Lease - ScopeId: $($DhcpScopeSelection.ScopeId) ScopeName: $($DhcpScopeSelection.Name) - v.$Version - $(Get-Date -Format "dddd, MMMM dd, yyyy  h:mm:ss tt")"
            $DhcpLeaseResultsSelection | Where-Object HostName -notlike '----------*' | Out-TableString -Wrap

            if ($DhcpLeaseResultsSelection.HostName -contains '---------- Refresh DHCP Leases ----------') {
                Invoke-DhcpPrompt -ComputerName $ComputerName -DhcpServer $DhcpServer -DhcpScopeId $DhcpScopeSelection.ScopeId -DhcpScopeName $DhcpScopeSelection.Name -Credential $Credential
            }
            elseif ($DhcpLeaseResultsSelection.HostName -contains '---------- DHCP Scope Selection ----------') {
                Invoke-DhcpPrompt -ComputerName $ComputerName -DhcpServer $DhcpServer -DhcpScopes $DhcpScopeResults -Credential $Credential
            }
        }
    }
    catch {
        Write-Warning $_.Exception.Message
    }
}



# START RUNNING the PROGRAM
VCAOpsPortal