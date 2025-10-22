#Harold.Kammermeyer@vca.com
#Requires -Modules ActiveDirectory
#Requires -Version 3

. "$PSScriptRoot\Convert-VcaAU.ps1"

function Get-VCAHeadCount {
    #Ver. 230516
    #Added support for new 20#### & 21#### prefix.
    #Ver. 200827
    #Results now exclude any CNF user accounts.
    #Empty/No OU results now return as 0 instead of 1.
    #Ver. 200430
    #Added support for H6000; unknown whether users will be added to OU. (Located under standard VCA OU)
    #Ver. 190430
    #Added support for Pet Partners (H5000) hospitals.
    #Ver. 190108
    #Fixed usercount when ou contained only 1 member.
    #Trimstart was causing issues with lists of sites.
    #Ver. 181115
    #Site users are now included in output.
    #Ver. 181003
    #Added -Full flag for grabbing all AD OU User count.

    [CmdletBinding()]
    param(
        [parameter(
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            Position = 0,
            ParameterSetName = 'QueryEach')]
        [alias('ComputerName', 'Name')]
        [string[]]$AU,

        [parameter(ParameterSetName = 'QueryAll')]
        [switch]$Full
    )
    begin {
        if ($Full.IsPresent) {
            Get-ADUser -Filter "Name -like '*' -and Name -notlike '*CNF:*'" -Properties CN |
                Select-Object -Property @{n = 'ParentContainer'; e = { $_.Distinguishedname -replace "CN=$($_.cn)," } } |
                Group-Object -Property ParentContainer | Select-Object -Property Name, @{n = 'UserCount'; e = { $_.Count } }
            return
        }

        $AU = Convert-VcaAU -AU $AU -Suffix ''
    }
    process {
        foreach ($AU_Item in $AU) {
            Clear-Variable -Name SiteUsers, ErrorMessage -ErrorAction Ignore

            try {
                $ADUserParams = @{
                    Properties  = @(
                        'City'
                        'State'
                        'StreetAddress'
                        'PostalCode'
                        'Department'
                        'EmailAddress'
                        'Office'
                        'OfficePhone'
                        'Title'
                        'Created'
                        'extensionAttribute3'
                    )
                    ErrorAction = 'Stop'
                }
                $auNumber = ($AU_Item -replace '^h', '').PadLeft(4, '0')
                $groupName = 'h' + $auNumber
                $SiteUsers = switch -regex ($AU_Item) {
                    '^h[0-9]{3}$|^h[0-3,6][0-9]{3}$|^h20[0-9]{4}$' { 
                        Get-ADGroupMember -Identity $groupName -Recursive | Get-ADUser @ADUserParams | Where-Object { $_.Name -notlike '*CNF:*' }; break 
                    } #VCA
                    '^h4[0-9]{3}$' { 
                        Get-ADGroupMember -Identity $groupName -Recursive | Get-ADUser @ADUserParams | Where-Object { $_.Name -notlike '*CNF:*' }; break 
                    }                        #CAPNA
                    '^h5[0-9]{3}$' { 
                        Get-ADGroupMember -Identity $groupName -Recursive | Get-ADUser @ADUserParams | Where-Object { $_.Name -notlike '*CNF:*' }; break 
                    }                 #Pet Partners
                    '^h8[0-9]{4}$|^h21[0-9]{4}$' { 
                        Get-ADGroupMember -Identity $groupName -Recursive | Get-ADUser @ADUserParams | Where-Object { $_.Name -notlike '*CNF:*' } 
                    }                   #VCA Canada
                    default { $ErrorMessage = 'Invalid AU format' }
                }
            }
            catch {
                $ErrorMessage = $_.Exception.Message
            }
            [pscustomobject]@{
                AU        = $AU_Item -replace '^h0',''
                UserCount = ($SiteUsers | Measure-Object).Count
                Users     = $SiteUsers
                Error     = $ErrorMessage
            }
        } #foreach
    } #process
    end {
        #Intentionally left blank
    }
} #function