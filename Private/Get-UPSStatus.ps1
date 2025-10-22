#Harold.Kammermeyer@vca.com
#v.191029
#Requires -Version 3

workflow Get-UpsStatus {
    [CmdletBinding()]
    param(
        [string[]]
        $UPSs
    )
    foreach -parallel ($ups in $UPSs) {
        inlinescript {
            try {
                if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type) {
                    $certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
                    Add-Type $certCallback
                }
                [ServerCertificateValidationCallback]::Ignore()

                $MgmtCard = (Invoke-Webrequest -Uri "http://$using:ups" -ErrorAction Stop).ParsedHtml.Title
            }
            catch {
                $ErrorMessage = $Error[0].Exception.Message
            }
            # G2/G4
            if ($MgmtCard -like '*UPS Network Module') {
                $Headers = @{'Authorization' = 'Basic YWRtaW46YWRtaW4=' }

                $UpsAbout = Invoke-WebRequest -Uri "http://$using:ups/ups_propAbout.htm"
                $UpsAbout2 = (($UpsAbout.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine0' }).innertext) -split (' :')
                $UpsAbout3 = (($UpsAbout.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine1' }).innertext) -split (' :')

                $UpsStatus = Invoke-WebRequest -Uri "http://$using:ups/ups_propStatus.htm"
                $UpsStatus1 = (($UpsStatus.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine0' }).innertext) -split (' :') -split '\n'
                $UpsStatus2 = (($UpsStatus.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine1' }).innertext) -split (' :') -split (': ') -replace '  ', '' -split '\n'

                $UpsAlarms = Invoke-WebRequest -Uri "http://$using:ups/ups_propAlarms.htm"
                $UpsAlarmsP = (($UpsAlarms.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine1' }).innertext)

                Invoke-WebRequest -Uri "http://$using:ups/pas_mgr.htm" | Out-Null
                $TrapCommunity = Invoke-WebRequest -Uri "http://$using:ups/pas_mgr.htm" -Headers $Headers
                $TrapCommunity2 = ($TrapCommunity.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.className -eq 'listLine1' }).innertext

                $SnmpCommunity = Invoke-WebRequest -Uri "http://$using:ups/set_snmp.htm" -Headers $Headers
            }
            # XR/G1
            elseif ($MgmtCard -eq 'HP UPS Management Module') {
                $UpsAbout = Invoke-WebRequest -Uri "http://$using:ups/data_ident.htm?tabID=0"
                $UpsAbout2 = $UpsAbout.Content -replace 'r\d{1,2}c\d\^', '' -split '\^txt\|'
                $UpsAbout = Invoke-WebRequest -Uri "http://$using:ups/data_ident.htm?tabID=1" 
                $UpsAbout3 = $UpsAbout.Content -replace 'r\d{1,2}c\d\^', '' -split '\^txt\|'
                $UpsStatus = Invoke-WebRequest -Uri "http://$using:ups/data_param.htm?tabID=0"
                $UpsStatusP = $UpsStatus.Content -replace 's\dr\d{1,2}c\d{1,2}\^', '' -replace '1\^icon\|', '' -replace '\d\^icon\|', '' -split '\^txt\|'
                $UpsAlarms = Invoke-WebRequest -Uri "http://$using:ups/data_alarms.htm?tabID=0"
                $UpsAlarmsP = $UpsAlarms.Content -replace 'r\d{1,2}c\d{1,2}\^', '' -replace '1\^icon\|', '' -replace '\d\^icon\|', '' -split '\^txt\|'

                Invoke-WebRequest -Uri "http://$using:ups/Forms/index_1" -Method 'POST' -Body 'userName=admin&password=admin&language=1&dummy3=Sign+In&cleanup=&ErrorMsg=0' -SessionVariable UpsSession -TimeoutSec 10 | Out-Null
                $TrapCommunity = Invoke-WebRequest -Uri "http://$using:ups/setup_NM_trapReceivers.htm" -WebSession $UpsSession -TimeoutSec 10
                $SnmpCommunity = Invoke-WebRequest -Uri "http://$using:ups/setup_NM_snmpManagers.htm" -WebSession $UpsSession -TimeoutSec 10
            }
            # APC
            elseif ($MgmtCard -like '*Log On*') {
                $UpsWebReference = (Invoke-WebRequest -Uri "http://$using:ups/Forms/login1" -Method "POST" -Headers @{"Cache-Control" = "max-age=0"; "Upgrade-Insecure-Requests" = "1"; "Content-Type" = "application/x-www-form-urlencoded"; "Accept-Encoding" = "gzip, deflate"; "Accept-Language" = "en-US,en;q=0.9" } -Body "prefLanguage=00000000&login_username=admin&login_password=admin&submit=Log+On" -SessionVariable UpsSession).Forms.Action -replace '/NMC/', '' -replace '/Forms/pghdnonav1', ''
                $UpsAbout = Invoke-WebRequest "http://$using:ups/NMC/$UpsWebReference/ulabout.htm" -WebSession $UpsSession
                $UpsAbout2 = ($UpsAbout.ParsedHtml.getElementsByTagName('tr') | Where-Object { $_.innerHTML -like '*dataName*' }).innerText[-9..-1] -split (':')
            }
            elseif ($MgmtCard -eq '') {
                $G5Token = Invoke-RestMethod -Uri "https://$using:ups/rest/mbdetnrs/1.0/oauth2/token" -Method "POST" -ContentType "application/json;charset=UTF-8" -Body "{`"username`":`"admin`",`"password`":`"#C3r3bru5!`",`"grant_type`":`"password`",`"scope`":`"GUIAccess`"}"
                $G5Results = Invoke-RestMethod -Uri "https://$using:ups/rest/mbdetnrs/1.0/powerDistributions/1?`$expand=5" -Headers @{"Accept" = "application/json, text/plain, */*"; "Authorization" = "$($G5Token.token_type) $($G5Token.access_token)"; "Cookie" = "eaton_token=$($G5Token.access_token)" }
            }
            if ($MgmtCard -like '*UPS Network Module') {
                [pscustomobject]@{
                    UPS                         = $using:ups
                    'UPS Name'                  = $(if ($UpsAbout2) { $UpsAbout2[1] })
                    'UPS Serial Number'         = $(if ($UpsAbout2) { $UpsAbout2[3] })
                    'Card Name'                 = $MgmtCard
                    'Card Firmware revision'    = $(if ($UpsAbout2) { $UpsAbout2[7] })
                    'Card Serial Number'        = $(if ($UpsAbout2) { $UpsAbout2[11] })
                    'Card Ethernet Speed'       = $(if ($UpsAbout2) { $UpsAbout2[13] })
                    'UPS Part Number'           = $(if ($UpsAbout3) { $UpsAbout3[1] })
                    'UPS Firmware Revision'     = $(if ($UpsAbout3) { $UpsAbout3[3] })
                    'Card Part Number'          = $(if ($UpsAbout3) { $UpsAbout3[5] })
                    'Card Hardware Revision'    = $(if ($UpsAbout3) { $UpsAbout3[7] })
                    'Card Ethernet Mac Address' = $(if ($UpsAbout3) { $UpsAbout3[9] })
                    'Power Source'              = $(if ($UpsStatus1) { $UpsStatus1[([array]::IndexOf($UpsStatus1, 'Power Source') + 2)] })
                    'Output load level'         = $(if ($UpsStatus1) { $UpsStatus1[([array]::IndexOf($UpsStatus1, 'Output load level') + 2)] })
                    'Battery Capacity'          = $(if ($UpsStatus1) { $UpsStatus1[([array]::IndexOf($UpsStatus1, 'Battery Capacity') + 2)] })
                    'Battery Status'            = $(if ($UpsStatus1) { $UpsStatus1[([array]::IndexOf($UpsStatus1, 'Battery Status'))] })
                    'Entire UPS Master'         = $(if ($UpsStatus2) { $UpsStatus2[([array]::IndexOf($UpsStatus2, 'Entire UPS Master') + 3)] })
                    'Load Segment 1'            = $(if ($UpsStatus2) { $UpsStatus2[([array]::IndexOf($UpsStatus2, 'Load Segment 1') + 1)] })
                    'Load Segment 2'            = $(if ($UpsStatus2) { $UpsStatus2[([array]::IndexOf($UpsStatus2, 'Load Segment 2') + 1)] })
                    'Remaining backup time'     = $(if ($UpsStatus2) { $UpsStatus2[([array]::IndexOf($UpsStatus2, 'Remaining backup time') + 1)] })
                    'Alarms'                    = $(if ($UpsAlarmsP) { $UpsAlarmsP })
                    'TrapCommunity'             = ''
                    'TrapIP'                    = $TrapCommunity2
                    'SnmpCommunity'             = $($SnmpCommunity.InputFields | Where-Object { $_.Name -eq 'ChangeCommunityReadOnly' }).Value
                    'SnmpIP '                   = ''
                    Error                       = $ErrorMessage
                }
            }
            elseif ($MgmtCard -eq 'HP UPS Management Module') {
                [pscustomobject]@{
                    UPS                         = $using:ups
                    'UPS Name'                  = $(if ($UpsAbout3) { $UpsAbout3[1] })
                    'UPS Serial Number'         = $(if ($UpsAbout3) { $UpsAbout3[9] })
                    'Card Name'                 = $MgmtCard
                    'Card Firmware revision'    = $(if ($UpsAbout2) { $UpsAbout2[19] })
                    'Card Serial Number'        = $(if ($UpsAbout2) { $UpsAbout2[17] })
                    'Card Ethernet Speed'       = ''
                    'UPS Part Number'           = $(if ($UpsAbout3) { $UpsAbout3[3] })
                    'UPS Firmware Revision'     = $(if ($UpsAbout3) { $UpsAbout3[7] })
                    'Card Part Number'          = ''
                    'Card Hardware Revision'    = $(if ($UpsAbout2) { $UpsAbout2[21] })
                    'Card Ethernet Mac Address' = $(if ($UpsAbout2) { $UpsAbout2[9] })
                    'Power Source'              = $(if (($UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Source'))]) -ge 0) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Source') + 1)] })
                    'Output load level'         = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Load') + 1)] })
                    'Battery Capacity'          = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Battery Capacity') + 1)] })
                    'Battery Status'            = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Battery Status') + 1)] })
                    'Entire UPS Master'         = ''
                    'Load Segment 1'            = ''
                    'Load Segment 2'            = ''
                    'Remaining backup time'     = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Run Time Remaining') + 1)] })
                    'Alarms'                    = $(if (($UpsAlarmsP) -and ($UpsAlarmsP -match '\d\d\/\d\d\/\d\d\d\d')) { $UpsAlarmsP[0..([array]::IndexOf($UpsAlarmsP, @($UpsAlarmsP | Where-Object { $_ -match '\d\d\:\d\d\:\d\d' })[-1]))] })
                    #'Battery Installed Date'      = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Battery Installed Date') + 1)] })
                    #'Battery Voltage'             = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Battery Voltage') + 1)] })
                    #'High Voltage Transfer Point' = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'High Voltage Transfer Point') + 1)] })
                    #'Low Voltage Transfer Point'  = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Low Voltage Transfer Point') + 1)] })
                    #'UPS Temperature'             = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'UPS Temperature') + 1)] -replace ' &deg;'})
                    #'Input Current'               = $(if (($UpsStatusP[([array]::IndexOf($UpsStatusP, 'Input Current'))]) -ge 0) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Input Current') + 1)] })
                    #'Input Frequency'             = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Input Frequency') + 1)] })
                    #'Input Voltage'               = $(if (($UpsStatusP[([array]::IndexOf($UpsStatusP, 'Input Voltage'))]) -ge 0) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Input Voltage') + 1)] })
                    #'Output Current'              = $(if (($UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Current'))]) -ge 0) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Current') + 1)] })
                    #'Output Frequency'            = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Frequency') + 1)] })
                    #'Output Power'                = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Power') + 1)] })
                    #'Output Voltage'              = $(if (($UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Voltage'))]) -ge 0) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Output Voltage') + 1)] })
                    #'Last Self Test Date'         = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Last Self Test Date') + 1)] })
                    #'Test Results Summary'        = $(if ($UpsStatusP) { $UpsStatusP[([array]::IndexOf($UpsStatusP, 'Test Results Summary') + 1)] })
                    'TrapCommunity'             = $($TrapCommunity.InputFields | Where-Object { $_.Name -eq 'community1' }).Value
                    'TrapIP'                    = $($TrapCommunity.InputFields | Where-Object { $_.Name -eq 'trapRec1_IP' }).Value
                    'SnmpCommunity'             = $($SnmpCommunity.InputFields | Where-Object { $_.Name -eq 'readStr1' }).Value
                    'SnmpIP'                    = $($SnmpCommunity.InputFields | Where-Object { $_.Name -eq 'snmpMan1_IP' }).Value
                    Error                       = $ErrorMessage
                }
            }
            elseif ($MgmtCard -like '*Log on*') {

                <#
                $UpsAbout2 =
                Model:Smart-UPS X 3000
                SKU:SMX3000LVNC
                Serial Number:AS1545343691 
                Firmware Revision:UPS 07.4 (ID1003) 
                Manufacture Date:11/08/2015
                Apparent Power Rating:2880 VA
                Real Power Rating:2700 W
                Internal Battery SKU:APCRBC143
                External Battery SKU:APCRBC143
                #>

                [pscustomobject]@{
                    UPS                         = $using:ups
                    'UPS Name'                  = $(if ($UpsAbout2) { $UpsAbout2[([array]::IndexOf($UpsAbout2, 'Model') + 1)] })
                    'UPS Serial Number'         = $(if ($UpsAbout2) { $UpsAbout2[([array]::IndexOf($UpsAbout2, 'Serial Number') + 1)] })
                    'Card Name'                 = $UpsAbout.ParsedHtml.Title
                    'Card Firmware revision'    = ''
                    'Card Serial Number'        = ''
                    'Card Ethernet Speed'       = ''
                    'UPS Part Number'           = $(if ($UpsAbout2) { $UpsAbout2[([array]::IndexOf($UpsAbout2, 'SKU') + 1)] })
                    'UPS Firmware Revision'     = $(if ($UpsAbout2) { $UpsAbout2[([array]::IndexOf($UpsAbout2, 'Firmware Revision') + 1)] })
                    'Card Part Number'          = ''
                    'Card Hardware Revision'    = ''
                    'Card Ethernet Mac Address' = ''
                    'Power Source'              = ''
                    'Output load level'         = ''
                    'Battery Capacity'          = ''
                    'Battery Status'            = ''
                    'Entire UPS Master'         = ''
                    'Load Segment 1'            = ''
                    'Load Segment 2'            = ''
                    'Remaining backup time'     = ''
                    'Alarms'                    = ''
                    'TrapCommunity'             = ''
                    'TrapIP'                    = ''
                    'SnmpCommunity'             = ''
                    'SnmpIP'                    = ''
                    Error                       = $ErrorMessage
                }
            }
            elseif ($MgmtCard -eq '') {
                [pscustomobject]@{
                    UPS                         = $using:ups
                    'UPS Name'                  = $G5Results.identification.name
                    'UPS Serial Number'         = $G5Results.identification.serialNumber
                    'Card Name'                 = $MgmtCard
                    'Card Firmware revision'    = ''
                    'Card Serial Number'        = ''
                    'Card Ethernet Speed'       = ''
                    'UPS Part Number'           = $G5Results.identification.partNumber
                    'UPS Firmware Revision'     = $G5Results.identification.firmwareVersion
                    'Card Part Number'          = ''
                    'Card Hardware Revision'    = ''
                    'Card Ethernet Mac Address' = ''
                    'Power Source'              = ''
                    'Output load level'         = $G5Results.output.phases.measures.percentLoad
                    'Battery Capacity'          = $G5Results.batteries.measures.remainingChargeCapacity
                    'Battery Status'            = ''
                    'Entire UPS Master'         = ''
                    'Load Segment 1'            = ''
                    'Load Segment 2'            = ''
                    'Remaining backup time'     = $G5Results.batteries.measures.remainingTime
                    'Alarms'                    = ''
                    'TrapCommunity'             = ''
                    'TrapIP'                    = ''
                    'SnmpCommunity'             = ''
                    'SnmpIP'                    = ''
                    Error                       = $ErrorMessage
                }
            }
            else {
                [pscustomobject]@{
                    UPS                         = $using:ups
                    'UPS Name'                  = ''
                    'UPS Serial Number'         = ''
                    'Card Name'                 = $MgmtCard
                    'Card Firmware revision'    = ''
                    'Card Serial Number'        = ''
                    'Card Ethernet Speed'       = ''
                    'UPS Part Number'           = ''
                    'UPS Firmware Revision'     = ''
                    'Card Part Number'          = ''
                    'Card Hardware Revision'    = ''
                    'Card Ethernet Mac Address' = ''
                    'Power Source'              = ''
                    'Output load level'         = ''
                    'Battery Capacity'          = ''
                    'Battery Status'            = ''
                    'Entire UPS Master'         = ''
                    'Load Segment 1'            = ''
                    'Load Segment 2'            = ''
                    'Remaining backup time'     = ''
                    'Alarms'                    = ''
                    'TrapCommunity'             = ''
                    'TrapIP'                    = ''
                    'SnmpCommunity'             = ''
                    'SnmpIP'                    = ''
                    Error                       = $ErrorMessage
                }
            } #else
        } #inlinescript
    } #foreach
} #workflow