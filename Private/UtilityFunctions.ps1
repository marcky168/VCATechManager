# Consolidated Utility Functions
# Contains: Convert-VcaAU, Copy-ToPSSession, Kill-SparkyShell, Remove-BakRegistry, Update-Changelog

function Convert-VcaAU {
    #Ver. 181211
     #Improved -ilo switch to work correctly with clustered naming convention.
     #Fixed issue with RetainSuffixNumber when a FQDN was provided.
    #Ver. 181115
     #All text is convereted to lowercase to accurately remove duplicates.
     #Added FQDN switch to add domain name to host.
     #Added Domain parameter to be used with -FQDN switch. Defaults to 'vcaantech.com'
     #Added ilo switch to auto-fill suffix.
    #Ver. 181021
     #Added esxihost & db switch.
    #Ver. 181003
     #Added util switch to auto-fill suffix.
     #Added strip switch for extracting AU number.
    #Ver. 180830
     #Added quser switch for quser output.
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0)]
        [alias('ComputerName', 'Name')]
        [string[]]$AU,
        [string]$Prefix = 'h',
        [string]$Suffix = '-ns',
        [switch]$Clipboard,
        [switch]$NoLeadingZeros,
        [switch]$RetainSuffixNumber,
        [switch]$Quser,
        [switch]$EsxiHost,
        [switch]$Database,
        [switch]$Util,
        [switch]$Ilo,
        [switch]$Strip,
        [switch]$FQDN,
        [string]$Domain = 'vcaantech.com'
    )
    begin {
        # Process clipboard items if clipboard switch was used.
        if ($Clipboard.IsPresent) {
            $AU = Get-Clipboard
        }
        # Remove empty lines
        $AU = $AU | Where-Object { $PSItem }
        # Remove duplicates
        $AU = $AU.ToLower() | Select-Object -Unique

        if ($Database.IsPresent) {
            $Suffix = '-db'
        }
        if ($EsxiHost.IsPresent) {
            $Suffix = '-vm'
        }
        if ($Util.IsPresent) {
            $Suffix = '-util'
        }
        if ($Ilo.IsPresent) {
            $Suffix = '-ilo'
            $RetainSuffixNumber = $true
        }
        if ($FQDN.IsPresent) {
            $Suffix += ".$Domain"
        }
        if ($Strip.IsPresent) {
            $Prefix, $Suffix = ''
            $NoLeadingZeros = $true
        }
    }
    process {
        foreach ($AU_Item in $AU) {
            if ($AU_SuffixNumber) { Clear-Variable -Name AU_SuffixNumber }
            if ($AU_ItemStripped) { Clear-Variable -Name AU_ItemStripped }

            # If input item is numbers only then store in $AU_ItemStripped and skip number extraction.
            if ($AU_Item -match '^[0-9]+$') {
                $AU_ItemStripped = $AU_Item
            }
            # Extract numerical AU # from string starting with case insensitive 'h' or 'au.'
            elseif ($AU_Item -match '^((?i)h|au)[0-9]{2,5}') {
                $AU_ItemStripped = ($AU_Item -replace ('^((?i)h|au)', '') -split '-')[0]

                # Extract suffix count, e.g. -ns01, -ups02
                If ($RetainSuffixNumber.IsPresent) {
                    if ($AU_Item -match '-[a-z]+[0-9]{1,2}') {
                        $AU_SuffixNumber = "$(($AU_Item -split '-')[1] -replace '[^0-9]+')"
                    }
                }
            }
            # Extract numerical AU # from string starting with numbers and leading up to a hyphen.
            elseif ($AU_Item -match '^[0-9]{2,5}-') {
                $AU_ItemStripped = ($AU_Item -split '-')[0]
            }

            # Perform output if AU number format match.
            if ($AU_ItemStripped) {
                # Format output prefix and suffix; fill in leading zeros.
                if ((-not $NoLeadingZeros.IsPresent) -and (-not $Quser.IsPresent)) {
                    if (-not $Ilo.IsPresent) {
                        "$Prefix{0}$AU_ItemStripped$Suffix$AU_SuffixNumber" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                    }
                    else {
                        "$Prefix{0}$AU_ItemStripped-vm$AU_SuffixNumber$Suffix" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                    }
                }
                # Format for quser
                elseif ($Quser.IsPresent) {
                    "quser /server:$Prefix{0}$AU_ItemStripped$Suffix$AU_SuffixNumber" -f ('0' * [math]::max(0, (4 - $AU_ItemStripped.length)))
                }
                # Remove leading zeros
                else {
                    $Prefix + $AU_ItemStripped.TrimStart('0') + $Suffix + $AU_SuffixNumber
                }
            }
            elseif ($AU_Item -like 'hmtprod-*') {
                ($AU_Item -split '-')[0]
            }
        } #foreach
    } #process
    end {
        # Output a return after conversion output when quser switch is specified.
        if ($Quser.IsPresent) { Write-Output "`r" }
    }
} #function

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

function Kill-SparkyShell {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AU
    )

    try {
        # Query and stop the SparkyShell process (adjust process name if needed)
        $processes = Get-Process -Name "SparkyShell" -ErrorAction SilentlyContinue
        if ($processes) {
            $processes | Stop-Process -Force
            Write-Log "Successfully killed SparkyShell processes for AU $AU."
        } else {
            Write-Log "No SparkyShell processes found for AU $AU."
        }
    } catch {
        Write-Host "Error in Kill-SparkyShell: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error killing SparkyShell for AU $($AU): $($_.Exception.Message)"
        throw
    }
}

##############################################################################################
##    Script to delete *.bak profile key from registry / Remove temp profiles from system
##    Author: Lokesh Agarwal
##    Input : servers parameter (Contains Servers name)
##############################################################################################
function Remove-BakRegistry {
	param(
		[string[]]$servers
	)

	Foreach ($server in $servers) {
		##connect with registry of remote machine
		$baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Localmachine", "$server")

		##set registry path
		$key = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileList", $true)

		## get all profile name
		$profilereg = $key.GetSubKeyNames()
		$profileregcount = $profilereg.count

		while ($profileregcount -ne 0) {
			## check for bak profiles

			if ($profilereg[$profileregcount - 1] -like "*.bak") {
				$bakname = $profilereg[$profileregcount - 1]

				$baknamefinal = $bakname.Split(".")[0]

				## Delete bak profile
			 $key.DeleteSubKeyTree("$bakname")


				##connect with profileGuid
				$keyGuid = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileGuid", $true)

				## get all profile Guid
				$Guidreg = $keyGuid.GetSubKeyNames()
				$Guidregcount = $Guidreg.count

				while ($Guidregcount -ne 0) {
					$bakname1 = $Guidreg[$Guidregcount - 1]

					$keyGuidTest = $baseKey.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion\ProfileGuid\$bakname1", $true)
					$KeyGuidSidValue = $keyGuidTest.GetValue("sidstring")
					$KeyGuidSidValue

					if ($baknamefinal -eq $KeyGuidSidValue) {
						## Delete Guid profile
						$keyGuid.DeleteSubKeyTree("$bakname1")
					}
					$Guidregcount = $Guidregcount - 1
				}


			}
			$profileregcount = $profileregcount - 1
		}
	}
} #function

function Update-Changelog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Version,

        [Parameter(Mandatory = $true)]
        [string]$Changes,

        [Parameter(Mandatory = $false)]
        [string]$Date = (Get-Date -Format "yyyy-MM-dd")
    )

    $changelogPath = Join-Path $PSScriptRoot "..\Marc-Tools-Changelog.txt"

    # Read existing changelog if it exists
    $existingContent = if (Test-Path $changelogPath) {
        Get-Content $changelogPath -Raw
    } else {
        ""
    }

    # Create new entry
    $newEntry = "Version $Version - $Date`n$Changes`n`n"

    # Prepend new entry to existing content
    $updatedContent = $newEntry + $existingContent

    # Write back to file
    Set-Content -Path $changelogPath -Value $updatedContent -Encoding UTF8

    Write-Log "Changelog updated for version $Version."
}