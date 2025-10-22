#Harold.Kammermeyer@vca.com
#Requires -Version 3
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