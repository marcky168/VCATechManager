#Harold.Kammermeyer@vca.com
#Requires -Modules Posh-SSH

function Get-DriveStatus {
    [CmdletBinding()]
    param(
        [parameter(Mandatory, Position = 0)]
        [string[]]$ComputerName,
        [parameter(Mandatory)]
        [pscredential]$Credential,
        [string]$ExportPath = "$PSScriptRoot\Reports",
        [switch]$NoLog
    )

    $Version = '200117'

    if (-not $NoLog.IsPresent) {
        if (-not(Test-Path -Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath | Out-Null }
    }

    #Bash script
    $BASHScript = @'
echo "> hostname";
hostname;echo '';

echo "> date";
date;echo '';

echo "> uptime";
uptime;echo '';

echo "> esxcli hardware platform get";
esxcli hardware platform get;echo '';

echo "> vmware -vl";
vmware -vl;echo '';

echo "> esxcli software vib list | grep -E 'scsi-hpsa|nhpsa|smartpqi|ssacli'";
esxcli software vib list | grep -E 'scsi-hpsa|nhpsa|smartpqi|ssacli';echo '';

hpclissacli='./opt/smartstorageadmin/ssacli/bin/ssacli'
hpclihpssacli='./opt/hp/hpssacli/bin/hpssacli'
hpclihpacucli='./opt/hp/hpacucli/bin/hpacucli'

hpcliarg1='ctrl all show status';
hpcliarg2='ctrl all show config detail';

if [ -f "$hpclissacli" ]; then
    echo "> $hpclissacli $hpcliarg1";
    ./opt/smartstorageadmin/ssacli/bin/ssacli $hpcliarg1;
    echo "> $hpclissacli $hpcliarg2";
    ./opt/smartstorageadmin/ssacli/bin/ssacli $hpcliarg2;

elif [ -f "$hpclihpssacli" ]; then
    echo "> $hpclihpssacli $hpcliarg1";
    ./opt/hp/hpssacli/bin/hpssacli $hpcliarg1

    echo "> $hpclihpssacli $hpcliarg2";    
    ./opt/hp/hpssacli/bin/hpssacli $hpcliarg2;

elif [ -f "$hpclihpacucli" ]; then
    echo "> $hpclihpacucli $hpcliarg1";
    ./opt/hp/hpacucli/bin/hpacucli $hpcliarg1;
    echo "> $hpclihpacucli $hpcliarg2";
    ./opt/hp/hpacucli/bin/hpacucli $hpcliarg2;

else echo 'Could not locate ESXCLI tool.';
fi;
'@.Replace("`r", '')

    ###Variables
    $ComputerName = $ComputerName.Trim()

    foreach ($ComputerName_Item in $ComputerName) {
        
        $ExportFile = "$ComputerName_Item-esxcli_$((Get-Date -UFormat %Y-%m-%d_%H-%M-%S)).txt"
        Write-Host "`nChecking: $ComputerName_Item`n"

        try {
            $SSHSession = New-SSHSession -ComputerName $ComputerName_Item -Credential $Credential -Force -ErrorAction Stop -WarningAction SilentlyContinue
            $SSHOutput = Invoke-SSHCommand -Command $BASHScript -SSHSession $SSHSession | Select-Object -ExpandProperty Output
            Remove-SSHSession -SSHSession $SSHSession | Out-Null
        }
        catch {
            Write-Warning "$($Error[0].Exception.Message)`n"
            if ($Error[0].Exception.Message -eq 'No connection could be made because the target machine actively refused it') {
                if (-not $Credential) { $Credential = Get-ESXiCreds }
                if ($Credential) {
                    Enable-SSH -ComputerName $ComputerName -Credential $Credential | Out-String
                }
            }
        }

        if ($SSHOutput) {
            $HeaderOutput = "ESXCLI Check v.$Version`r`n"

            #Filter esxcli failures
            $PatternMatch = 'fail|rebuild|recovering|error|Temporarily Disabled|Permanently Disabled|Cache disabled|Parity Initialization'
            $PatternExclude = 'Spare Activation Mode|Rebuild Priority|Unrecoverable Media|Initialization Completed'

            $SSHFailures = $SSHOutput | Select-String -Pattern $PatternMatch | Select-String -Pattern $PatternExclude -NotMatch
            
            #Summarize failures
            if (-not $SSHFailures) {
                $HeaderOutput += "No failures found. Please review log.`r`n"
            }
            else {
                $HeaderOutput += "Failures found:"
                $HeaderOutput = Write-Output $HeaderOutput $SSHFailures
            }
            #Draw Horizontal line
            $HeaderOutput += "`r`n$('#'*90)`r`n"
            #Draw arrows to failures
            $SSHOutput | Select-String -Pattern $PatternMatch | Select-String -Pattern $PatternExclude -NotMatch | Select-Object -Unique |
                ForEach-Object { $SSHOutput = $SSHOutput -replace [regex]::escape($_), ("$_ <$('-'*35)") }

            #Save SSH Results to file
            if (-not $NoLog.IsPresent) {
                try {
                    Write-Output $HeaderOutput $SSHOutput | Out-File -FilePath "$ExportPath\$ExportFile" -Encoding ascii -ErrorAction Stop
                    Start-Process -FilePath "$ExportPath\$ExportFile"
                }
                catch {
                    Write-Warning $Error[0].Exception.Message
                    
                    $ExportPath = "$($PSScriptRoot | Split-Path -Parent)\Reports\"
                    if (-not(Test-Path -Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath | Out-Null }
                    Write-Output $HeaderOutput $SSHOutput | Out-File -FilePath "$ExportPath\$ExportFile" -Encoding ascii
                    Start-Process -FilePath "$ExportPath\$ExportFile"
                }

                if (Test-Path -Path "$ExportPath\$ExportFile") {
                    Write-Host "`nReport saved to:"
                    Write-Host "$ExportPath\$ExportFile"
                }
                Write-Output $SSHFailures
            }
            else {
                if ($PSCmdlet.MyInvocation.BoundParameters['Debug'].IsPresent) {
                    $global:EsxcliDebug = $SSHOutput
                }

                #Write to screen
                Write-Output $HeaderOutput $SSHOutput

                #Draw Horizontal line
                Write-Output "`r"
                Write-Output "$('#'*90)`r"
                Write-Output ' Summary results may be innacurate or incomplete, please review full log.'
                Write-Output "$('#'*90)`r`n"
                Write-Output $ComputerName_Item `n
                Write-Output $SSHFailures `n

                #Get match context range
                if (@($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$').count -eq 1) {
                    $DriveContext = ($SSHOutput | Select-String -Pattern 'LD Acceleration Method: ').LineNumber -
                    ($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$').LineNumber

                    if ($DriveContext) {
                        Write-Output ($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$' -Context 0, $DriveContext) `n
                        Write-Output $SSHOutput | Select-String -Pattern 'Physical Drives$' -Context 0, 8
                        <#
                        $DriveContext2 = ($SSHOutput | Select-String -Pattern 'Internal Drive Cage at Port 2I').LineNumber -
                        ($SSHOutput | Select-String -Pattern 'Physical Drives$')[0].LineNumber - 2
                        Write-Output ($SSHOutput | Select-String -Pattern 'Physical Drives$' -Context 0, $DriveContext2) `n

                        $DriveContext3 = ($SSHOutput | Select-String -Pattern 'Array: A$').LineNumber -
                        ($SSHOutput | Select-String -Pattern 'Physical Drives$')[1].LineNumber - 2
                        Write-Output ($SSHOutput | Select-String -Pattern 'Physical Drives$' -Context 0, $DriveContext3) `n
                        #>
                    }
                }
                elseif (@($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$').count -ge 2) {
                    $DriveContext = ($SSHOutput | Select-String -Pattern 'LD Acceleration Method: ')[0].LineNumber -
                    ($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$')[0].LineNumber

                    if ($DriveContext) {
                        Write-Output ($SSHOutput | Select-String -Pattern 'Array: [a-zA-Z]$' -Context 0, $DriveContext) `n

                        $DriveContext = ($SSHOutput | Select-String -Pattern 'Port Name: 1I').LineNumber -
                        ($SSHOutput | Select-String -Pattern 'Physical Drives$').LineNumber - 2
                        Write-Output ($SSHOutput | Select-String -Pattern 'Physical Drives$' -Context 0, $DriveContext) `n
                    }
                }
            }
        } #if ($SSHOutput)
    } #foreach
} #function