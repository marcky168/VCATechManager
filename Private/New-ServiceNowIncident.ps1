#Harold.Kammermeyer@vca.com

function New-ServiceNowIncident {
    param (
        [string]$ImpactedUser = "",
        [string]$ReportedBy = "",
        [string]$Category = "Inquiry/Help",
        [string]$CIName = "",
        [string]$Impact = "3 - Low",
        [string]$Urgency = "2 - Medium",
        [string]$ContactType = "Monitoring",
        [string]$AssignedGroup = "Operations",
        [string]$AssignedTo = "",
        [string]$ShortDescription = "",
        [string]$Description = "",
        [string]$SnowApiUri = "vca.service-now.com",
        [pscredential]$Credential
    )
    
    begin {
        if (-not $Credential) { $Credential = Get-Credential -Message "ServiceNow API Credentials:" -UserName 'API.Operations' }
        if ($SnowApiUri -notlike '*.service-now.com') { $SnowApiUri = "$SnowApiUri.service-now.com" }
        $SnowApiUri = $SnowApiUri -replace 'http://',''
    }
    process {
        if ($Credential) {
            $WWAPIUser = $Credential.GetNetworkCredential().UserName
            $WWAPIPW = $Credential.GetNetworkCredential().Password
    
            $Headers = @{ Authorization = "Basic " + [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$WWAPIUser`:$WWAPIPW")) }

            $ObjPayload = @{
                ImpactedUser     = $ImpactedUser #String email address of SNOW user:Default Guest
                ReportedBy       = $ReportedBy #String email address of SNOW user:Default Guest
    
                Category         = $Category #String accepted values:"hardware";"inquiry";"network";"software";"database";"security":Default "inquiry"
                CIName           = $CIName #String:Name of CI:Default "Hardware - Other"
                Impact           = $Impact #Numeric:Range 1 to 3:Default 3
                Urgency          = $Urgency #Numeric:Range 1 to 3:Default 3
            
                ContactType      = $ContactType #String: accepted values: "messenger";"email";"phone";"self-service";"monitoring";"voice mail";"walk-in/direct-contact":DEFAULT "monitoring"
                AssignedGroup    = $AssignedGroup #String:Name of SNOW Group:Default "Support Alerts"
                AssignedTo       = $AssignedTo #Nameof of a member of SNOW group:Default NULL if AssignedTo user is not a member of the AssignedGroup 
    
                ShortDescription = $ShortDescription #String: Default value "Short Description is MISSING"
                Description      = $Description #String: Default value "Description is MISSING"
            }
            $ObjJson = $ObjPayload | ConvertTo-Json
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $WWAPIResult = (Invoke-RestMethod -Uri "https://$SnowApiUri/api/vcaan/api_operations/open" -Method Post -Body $ObjJson -ContentType 'application/json' -Headers $Headers -ErrorAction Stop).result
        
                $WWAPIResultObj = [pscustomobject]@{
                    incident = $WWAPIResult.incident
                    sys_id   = $WWAPIResult.sys_id
                    url      = "https://$SnowApiUri/nav_to.do?uri=%2Fincident.do%3Fsys_id%3D$($WWAPIResult.sys_id)%26sysparm_stack%3D%26sysparm_view%3Ddefault%26sysparm_view_forced%3Dtrue"
                }
                Write-Output $WWAPIResultObj
            }
            catch {
                Write-Warning $Error[0].Exception.Message
            }
        }
    }
    end {
        if ($WWAPIResult) {
            if ($Credential -and (-not (Get-StoredCredential -Target vcasnowapi))) {
                New-StoredCredential -Credentials $Credential -Target vcasnowapi -Type Generic -Persist LocalMachine | Out-Null
                Set-Variable -Scope Global -Name SNOWAPICredential -Value $Credential
            }
        }
    }
}