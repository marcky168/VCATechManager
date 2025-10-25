# Consolidated ServiceNow Functions
# Contains: New-ServiceNowGUI, New-ServiceNowIncident

function New-ServiceNowGUI {
    param (
        [string]$AU,
        [string]$ImpactedUser = "",
        [string[]]$ImpactedUserList = "",
        [string]$ReportedBy = "",
        [string]$Category = "Hardware",
        [string]$CIName = "",
        [string]$Impact = "3 - Low",
        [string]$Urgency = "2 - Medium",
        [string]$ContactType = "Monitoring",
        [string]$AssignedGroup = "VCA Operations",
        [string]$AssignedTo = "",
        [string]$ShortDescription = "",
        [string]$Description = "",
        [string[]]$ComputerName,
        [pscredential]$Credential
    )

    Add-Type -AssemblyName presentationframework

    $InputXML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="Generate ServiceNow Incident" Height="459.96" Width="810.864" ResizeMode="CanMinimize">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <ComboBox x:Name="cboImpactedUser" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,10,0,0" IsEditable="True" ToolTip="Impacted user"/>
        <ComboBox x:Name="cboReportedBy" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,37,0,0" IsEditable="True" ToolTip="Reported by"/>
        <ComboBox x:Name="cboCategory" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,64,0,0" IsReadOnly="True" ToolTip="Category"/>
        <ComboBox x:Name="cboConfigurationItem" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,91,0,0" IsReadOnly="True" ToolTip="Configuration Item"/>
        <ComboBox x:Name="cboImpact" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,118,0,0" IsReadOnly="True" ToolTip="Impact"/>
        <ComboBox x:Name="cboUrgency" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="10,145,0,0" IsReadOnly="True" ToolTip="Urgency"/>
        <ComboBox x:Name="cboContactType" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="487,10,0,0" ToolTip="Contact Type"/>
        <ComboBox x:Name="cboAssignedGroup" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="487,37,0,0" IsReadOnly="True" ToolTip="Assigned Group"/>
        <ComboBox x:Name="cboAssignedTo" HorizontalAlignment="Left" VerticalAlignment="Top" Width="203" Margin="487,64,0,0" IsEditable="True" ToolTip="Assigned To"/>
        <TextBox x:Name="txtShortDescription" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="774" Margin="10,172,0,0" ToolTip="Short Description"/>
        <TextBox x:Name="txtDescription" HorizontalAlignment="Left" Height="214.269" TextWrapping="Wrap" VerticalAlignment="Top" Width="774" Margin="10,200,0,0" VerticalScrollBarVisibility="Visible" UseLayoutRounding="False" AcceptsReturn="True" ToolTip="Description"/>
        <Button x:Name="cmdGenerate" Content="_Generate" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="707,10,0,0"/>

    </Grid>
</Window>
"@

    [xml]$XAML = $InputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'

    #Read XAML
    $reader = (New-Object System.Xml.XmlNodeReader $XAML)
    try {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch {
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        throw
    }

    #===========================================================================
    # Load XAML Objects In PowerShell
    #===========================================================================

    $XAML.SelectNodes("//*[@Name]") | ForEach-Object {
        try {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop
        }
        catch { throw }
    }

    #===========================================================================
    # Use this space to add code to the various form elements in your GUI
    #===========================================================================

    # Get user's email address using .NET
    $ADQuery = New-Object System.DirectoryServices.DirectorySearcher
    $ADQuery.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
    $ADQuery.Filter = "(&(ObjectClass=User)(samAccountName=$env:USERNAME))"
    $ADQueryResult = $ADQuery.FindOne()

    $UserEmail = $($ADQueryResult.Properties.mail)

    #Impacted User
    @(
        ''
    ) +
    $ImpactedUserList | ForEach-Object {
        [void]$WPFcboImpactedUser.items.Add($PSItem)
    }
    $WPFcboImpactedUser.Text = "$ImpactedUser"

    #Reported By
    @(
        ''
        "$UserEmail"
    ) | ForEach-Object {
        [void]$WPFcboReportedBy.items.Add($PSItem)
    }
    $WPFcboReportedBy.Text = "$UserEmail"

    #Category
    @(
        ''
        'Hardware'
        'Enterprise Application'
        'Inquiry/Help'
        'Network'
        'Software'
        'Database'
        'PMS'
        'Telecom'
        'Security'
        '8x8 Phone System'
    ) | ForEach-Object {
        [void]$WPFcboCategory.items.Add($PSItem)
    }
    $WPFcboCategory.Text = "$Category"

    #Configuration Item
    @(
        ''
        'eFilm'
        'Router - ISP'
        'Server - Hospital'
        'UPS'
        'Windows OS'
    ) +
    $ComputerName | ForEach-Object {
        [void]$WPFcboConfigurationItem.items.Add($PSItem)
    }
    $WPFcboConfigurationItem.SelectedValue = "$CIName"
    #$WPFcboConfigurationItem.Text = "$CIName"

    #Impact
    @(
        '1 - High'
        '2 - Medium'
        '3 - Low'
    ) | ForEach-Object {
        [void]$WPFcboImpact.items.Add($PSItem)
    }
    $WPFcboImpact.Text = "$Impact"

    #Urgency
    @(
        '1 - High'
        '2 - Medium'
        '3 - Low'
    ) | ForEach-Object {
        [void]$WPFcboUrgency.items.Add($PSItem)
    }
    $WPFcboUrgency.Text = "$Urgency"

    #Contact Type
    @(
        'Chat'
        'Email'
        'Emergency Line'
        'Monitoring'
        'Phone'
        'Project'
        'Self-service'
        'Voice mail'
        'Walk-In/Direct-Contact'
    ) | ForEach-Object {
        [void]$WPFcboContactType.items.Add($PSItem)
    }
    $WPFcboContactType.Text = 'Monitoring'

    #AssignedGroup
    @(
        'VCA Operations'
        'VCA Administrative'
        'VCA Automation-Integration'
        'VCA Cognizant Service Desk'
        'VCA DataCenter'
        'VCA Engineering-Hospital'
        'VCA Engineering-Network'
        'VCA Engineering-Security'
        'VCA Engineering-Systems'
        'VCA Hospital Support'
        'VCA iLink SCOM Team'
        'VCA iLink BigFix Team'
        'VCA iLink Patching'
        'VCA SQL DBA'
        'VCA Telecom'
        'VCA WOOFware Support'
        'VCAC Support'
        'QOS NOC'
    ) | ForEach-Object {
        [void]$WPFcboAssignedGroup.items.Add($PSItem)
    }
    $WPFcboAssignedGroup.Text = 'VCA Operations'

    #Assigned To
    @(
        ''
    "$UserEmail"
    ) | ForEach-Object {
        [void]$WPFcboAssignedTo.items.Add($PSItem)
    }
    $WPFcboAssignedTo.Text = "$UserEmail"

    #Short Description
    $WPFtxtShortDescription.Text = "$ShortDescription"

    #Events
    $WPFcmdGenerate.Add_Click( {
            $MsgBoxInput = [System.Windows.MessageBox]::Show('Are you sure?', 'Generate ServiceNow Incident', 'YesNo', 'Question')
            switch ($MsgBoxInput) {
                'Yes' {
                    Clear-Variable -Name WWAPIResult -ErrorAction Ignore
                    $NewServiceNowParams = [ordered]@{
                        ImpactedUser     = $WPFcboImpactedUser.Text
                        ReportedBy       = $WPFcboReportedBy.Text
                        Category         = $WPFcboCategory.Text
                        CIName           = $WPFcboConfigurationItem.Text
                        Impact           = $WPFcboImpact.Text
                        Urgency          = $WPFcboUrgency.Text
                        ContactType      = $WPFcboContactType.Text
                        AssignedGroup    = $WPFcboAssignedGroup.Text
                        AssignedTo       = $WPFcboAssignedTo.Text
                        ShortDescription = $WPFtxtShortDescription.Text
                        Description      = $WPFtxtDescription.Text
                    }
                    Write-Host ($NewServiceNowParams | Out-String) -ForegroundColor Cyan
                    if ($Credential) { $NewServiceNowParams.Credential = $Credential }

                    $WWAPIResult = New-ServiceNowIncident @NewServiceNowParams

                    if ($WWAPIResult) {
                        #$Form.Close()
                        Write-Host "`rTicket Generated: $($WWAPIResult.incident) - $($WPFtxtShortDescription.Text)" -ForegroundColor Cyan
                        Write-Host "`n$($WWAPIResult.url)"
                        Start-Process $WWAPIResult.url
                        Write-Host ''
                    }
                }
                'No' {
                }
            }
        })
    #===========================================================================
    # Shows the form
    #===========================================================================
    $async = $Form.Dispatcher.InvokeAsync( {
            $Form.ShowDialog() | Out-Null
        })
    $async.Wait() | Out-Null
}

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