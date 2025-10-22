# Harold.Kammermeyer@vca.com
# Copy file via PSSession with path creation and optional hash verification.
function Copy-ToPSSession {
    [CmdletBinding()]
    param(
        [string[]]$ComputerName,
        [string[]]$Path,
        [string]$Destination,
        [switch]$VerifyHash,
        [string]$Algorithm = 'MD5',
        [System.Management.Automation.Runspaces.PSSession[]]$Session,
        [pscredential]$Credential
    )
    begin {
        if (-not (Test-Path -Path $Path)) { Write-Warning "[$Path] Does not exist."; break }

        $FullDestinationPath = "$Destination\$(Split-Path -Path $Path -Leaf)"

        # Create file hash if parameter was used
        if ($VerifyHash.IsPresent) { $Hash = (Get-FileHash -Path $Path -Algorithm $Algorithm).Hash }

        # Create pssession if it doesn't exist
        if (-not $Session) {
            foreach ($ComputerName_Item in $ComputerName) {
                try {
                    $Session = $Session + (New-PSSession -ComputerName $ComputerName_Item -Credential $Credential -ErrorAction Stop)
                }
                catch {
                    Write-Warning "[$ComputerName_Item] $($PSItem.Exception.Message)"
                }
            }
        }
    }
    process {
        foreach ($Session_Item in $Session) {
            # Check if file exists in destination folder
            if (-not (Invoke-Command -Session $Session { Test-Path -Path $using:FullDestinationPath })) {
                # Create destination path if it doesn't exist
                Invoke-Command -Session $Session {
                    if (-not (Test-Path -Path $using:Destination)) { New-Item -ItemType Directory -Path $using:Destination | Out-Null }
                }
                # Copy file to pssession
                try {
                    Write-Host "[$($Session_Item.ComputerName)] Copying $(Split-Path -Path $Path -Leaf)" -ForegroundColor Cyan
                    Copy-Item -Path $Path -Destination $Destination -ToSession $Session_Item -ErrorAction Stop
                }
                catch {
                    Write-Warning $_.Exception.Message
                }
            }
            # Verify hash
            if ($VerifyHash.IsPresent -and $Hash -ne '') {
                # Verify file exists
                $ScriptBlock = {
                    if (Test-Path -Path $using:FullDestinationPath) {
                        if ((Get-FileHash -Path $using:FullDestinationPath -Algorithm $using:Algorithm).Hash -ne $using:Hash) {
                            Write-Warning "Source Hash does not match destination."
                        }
                    }
                }
                Invoke-Command -Session $Session $ScriptBlock
            }
        }
    }
    end {
        if ($ComputerName -and $Session) { Remove-PSSession -Session $Session }
    }
}