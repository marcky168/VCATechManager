<# (c) Copyright 2015 Hewlett Packard Enterprise Development LP #>
Add-Type @'
public class AsyncPipeline
{
    public System.Management.Automation.PowerShell Pipeline ;
    public System.IAsyncResult AsyncResult ;
}
'@
function Create-ThreadPool
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Position=0,Mandatory=$true)][int]$PoolSize,
        [Parameter(Position=1,Mandatory=$False)][Switch]$MTA
    )
    
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $PoolSize)	
    
    If(!$MTA) { $pool.ApartmentState = 'STA' }
    
    $pool.Open()
    
    return $pool
}

function Start-ThreadScriptBlock
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Position=0,Mandatory=$True)]$ThreadPool,
        [Parameter(Position=1,Mandatory=$True)][ScriptBlock]$ScriptBlock,
        [Parameter(Position=2,Mandatory=$False)][Object[]]$Parameters
    )
    
    $Pipeline = [System.Management.Automation.PowerShell]::Create() 

	$Pipeline.RunspacePool = $ThreadPool
	    
    $Pipeline.AddScript($ScriptBlock) | Out-Null
    
    Foreach($Arg in $Parameters)
    {
        $Pipeline.AddArgument($Arg) | Out-Null
    }
    
	$AsyncResult = $Pipeline.BeginInvoke() 
	
	$Output = New-Object AsyncPipeline 
	
	$Output.Pipeline = $Pipeline
	$Output.AsyncResult = $AsyncResult
	
	$Output
}

function Get-ThreadPipelines
{
    [Cmdletbinding()]
    Param
    (
        [Parameter(Position=0,Mandatory=$True)][AsyncPipeline[]]$Pipelines,
		[Parameter(Position=1,Mandatory=$false)][Switch]$ShowProgress
    )
	
	# incrementing for Write-Progress
    $i = 1 
	
    foreach($Pipeline in $Pipelines)
    {
		try
		{
        	$Pipeline.Pipeline.EndInvoke($Pipeline.AsyncResult)

			If($Pipeline.Pipeline.Streams.Error)
			{
				Throw $Pipeline.Pipeline.Streams.Error
			}
        } catch {
			$_
		}
        $Pipeline.Pipeline.Dispose()
		$i++
    }
}


$ThreadPipes = @()
$poolsize = 1
$thispool = Create-ThreadPool $poolsize
$t = 
{
    Param($path)
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') 
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

    $mod = "HPERedfishCmdlets"
    $RegPath = 'HKCU:\Software\Hewlett-Packard\PowerShell\Modules\' + $mod
    $PropLicense = 'LicenseShown'
    $PropPath = 'InstallPath'
    $PropVersion = 'ModuleVersion'
    
    $alreadyInstalled = $false
    $sameInstalledVersion = $true

    $modInfoPath = $path + '\'+ $mod + '_aadc4b97-c04c-44c6-8d69-1ebc5b5ffcc8_ModuleInfo.xml'
    $xmlLocal = [xml](Get-Content -Path "$modInfoPath")
    $currVersionString = $xmlLocal.ModuleInfo.Version.ToString()

    if ($(Test-Path $RegPath) -and $(Get-ItemProperty -Path $RegPath -Name $PropLicense -ErrorAction SilentlyContinue).$PropLicense -eq 'True')
    {
        $alreadyInstalled = $True
        
        $oldVersionString = (Get-ItemProperty -Path $RegPath -Name $PropVersion).$PropVersion

        if($currVersionString -ne $oldVersionString)
        {
            $sameInstalledVersion = $false
        }
    }

    #check to see if the registry path and entry are present
    if ($alreadyInstalled -eq $false -or $sameInstalledVersion -eq $false) {
        #show the license and create the key
    
        ################################################
        
        $objForm = New-Object System.Windows.Forms.Form 
        $objForm.Text = $mod + ' License'
        $objForm.Size = New-Object System.Drawing.Size(762,660) 
        $objForm.StartPosition = 'CenterScreen'

        $objForm.KeyPreview = $True
        $x = 'Cancel'
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Size(75,582)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Agree'
        $OKButton.Add_Click({$global:x='Agree';$objForm.Close()})
        $objForm.Controls.Add($OKButton)

        $objLabel = New-Object System.Windows.Forms.Label
        $objLabel.Font = New-Object System.Drawing.Font('SansSerif',12)
        $objLabel.Location = New-Object System.Drawing.Size(10,20) 
        $objLabel.Size = New-Object System.Drawing.Size(700,30)
        $objLabel.Text = 'By using this module you agree to these license terms:'
        $objForm.Controls.Add($objLabel) 
        $link=''
        $objRTFTextBox = New-Object System.Windows.Forms.RichTextBox 
        $objRTFTextBox.Location = New-Object System.Drawing.Size(10,60) 
        $objRTFTextBox.Size = New-Object System.Drawing.Size(728,500)

        $eulapath = $path + '\' + $mod + 'License.rtf'

        $objRTFTextBox.LoadFile($eulapath)
        $objRTFTextBox.ReadOnly = $True
        $objRTFTextBox.DetectUrls = $True
        $objRTFTextBox.Add_LinkClicked({$global:link = $_.LinkText; start $_.LinkText})
        $objForm.Controls.Add($objRTFTextBox) 

        $objForm.Topmost = $True

        $size = New-Object System.Drawing.Size(0,0)
        #Home key grabs window size
        $objForm.Add_KeyDown({if ($_.KeyCode -eq 'Home') 
            {
                $global:size.Height = $global:objForm.Size.Height
                $global:size.Width = $global:objForm.Size.Width
                $_.Handled = $True
            }
        })

        $objForm.ControlBox = $false
        $objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $objForm.Add_Shown({$objForm.Activate()})
        [void] $objForm.ShowDialog()
        
        #create the path first
        $newpath = ''
        $pathitems = $RegPath.Split('\')
        foreach ($pathitem in $pathitems) {
            if ($pathitem -notmatch ':') {
                $newpath += '\' + $pathitem
                if (-not $(Test-Path $newpath)) {
                    New-Item -Path $newpath | Out-Null
                }
            }
            else {
                $newpath = $pathitem
            }
        }
        
        $modInfoPath = $path + '\'+ $mod + '_aadc4b97-c04c-44c6-8d69-1ebc5b5ffcc8_ModuleInfo.xml'
        $xmlLocal = [xml](Get-Content -Path "$modInfoPath")
        $currVersionString = $xmlLocal.ModuleInfo.Version.ToString()
        
        #create the property
        if($alreadyInstalled -eq $false)
        {
            New-ItemProperty -Path $RegPath -Name $PropVersion -PropertyType String -Value $currVersionString | Out-Null
            New-ItemProperty -Path $RegPath -Name $PropLicense -PropertyType String -Value $true | Out-Null
            New-ItemProperty -Path $RegPath -Name $PropPath -PropertyType String -Value $path | Out-Null
        }
        if($sameInstalledVersion -eq $false)
        {
            Set-ItemProperty -Path $RegPath -Name $PropVersion -Value $currVersionString | Out-Null
        }
        
    }
} 
		#end of $t scriptblock
            
$ThreadPipes += Start-ThreadScriptBlock -ThreadPool $thispool -ScriptBlock $t -Parameters $PSScriptRoot.ToString()

if ($VerbosePreference -eq 'Continue')
{
	$rstList = Get-ThreadPipelines -Pipelines $ThreadPipes -ShowProgress
}
else
{
	$rstList = Get-ThreadPipelines -Pipelines $ThreadPipes
}
$thispool.Close()
$thispool.Dispose()
# SIG # Begin signature block
# MIIkXwYJKoZIhvcNAQcCoIIkUDCCJEwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCoyRrAj3mfNTWB
# Qhk4OiZDzm7FhNqJaKFaBl5q15C31KCCHtQwggQUMIIC/KADAgECAgsEAAAAAAEv
# TuFS1zANBgkqhkiG9w0BAQUFADBXMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xv
# YmFsU2lnbiBudi1zYTEQMA4GA1UECxMHUm9vdCBDQTEbMBkGA1UEAxMSR2xvYmFs
# U2lnbiBSb290IENBMB4XDTExMDQxMzEwMDAwMFoXDTI4MDEyODEyMDAwMFowUjEL
# MAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMT
# H0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzIwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQCU72X4tVefoFMNNAbrCR+3Rxhqy/Bb5P8npTTR94ka
# v56xzRJBbmbUgaCFi2RaRi+ZoI13seK8XN0i12pn0LvoynTei08NsFLlkFvrRw7x
# 55+cC5BlPheWMEVybTmhFzbKuaCMG08IGfaBMa1hFqRi5rRAnsP8+5X2+7UulYGY
# 4O/F69gCWXh396rjUmtQkSnF/PfNk2XSYGEi8gb7Mt0WUfoO/Yow8BcJp7vzBK6r
# kOds33qp9O/EYidfb5ltOHSqEYva38cUTOmFsuzCfUomj+dWuqbgz5JTgHT0A+xo
# smC8hCAAgxuh7rR0BcEpjmLQR7H68FPMGPkuO/lwfrQlAgMBAAGjgeUwgeIwDgYD
# VR0PAQH/BAQDAgEGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFEbYPv/c
# 477/g+b0hZuw3WrWFKnBMEcGA1UdIARAMD4wPAYEVR0gADA0MDIGCCsGAQUFBwIB
# FiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAzBgNVHR8E
# LDAqMCigJqAkhiJodHRwOi8vY3JsLmdsb2JhbHNpZ24ubmV0L3Jvb3QuY3JsMB8G
# A1UdIwQYMBaAFGB7ZhpFDZfKiVAvfQTNNKj//P1LMA0GCSqGSIb3DQEBBQUAA4IB
# AQBOXlaQHka02Ukx87sXOSgbwhbd/UHcCQUEm2+yoprWmS5AmQBVteo/pSB204Y0
# 1BfMVTrHgu7vqLq82AafFVDfzRZ7UjoC1xka/a/weFzgS8UY3zokHtqsuKlYBAIH
# MNuwEl7+Mb7wBEj08HD4Ol5Wg889+w289MXtl5251NulJ4TjOJuLpzWGRCCkO22k
# aguhg/0o69rvKPbMiF37CjsAq+Ah6+IvNWwPjjRFl+ui95kzNX7Lmoq7RU3nP5/C
# 2Yr6ZbJux35l/+iS4SwxovewJzZIjyZvO+5Ndh95w+V/ljW8LQ7MAbCOf/9RgICn
# ktSzREZkjIdPFmMHMUtjsN/zMIIEnzCCA4egAwIBAgISESEGoIHTP9h65YJMwWtS
# CU4DMA0GCSqGSIb3DQEBBQUAMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9i
# YWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFtcGluZyBD
# QSAtIEcyMB4XDTE1MDIwMzAwMDAwMFoXDTI2MDMwMzAwMDAwMFowYDELMAkGA1UE
# BhMCU0cxHzAdBgNVBAoTFkdNTyBHbG9iYWxTaWduIFB0ZSBMdGQxMDAuBgNVBAMT
# J0dsb2JhbFNpZ24gVFNBIGZvciBNUyBBdXRoZW50aWNvZGUgLSBHMjCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBALAXrqLTtgQwVh5YD7HtVaTWVMvY9nM6
# 7F1eqyX9NqX6hMNhQMVGtVlSO0KiLl8TYhCpW+Zz1pIlsX0j4wazhzoOQ/DXAIlT
# ohExUihuXUByPPIJd6dJkpfUbJCgdqf9uNyznfIHYCxPWJgAa9MVVOD63f+ALF8Y
# ppj/1KvsoUVZsi5vYl3g2Rmsi1ecqCYr2RelENJHCBpwLDOLf2iAKrWhXWvdjQIC
# KQOqfDe7uylOPVOTs6b6j9JYkxVMuS2rgKOjJfuv9whksHpED1wQ119hN6pOa9PS
# UyWdgnP6LPlysKkZOSpQ+qnQPDrK6Fvv9V9R9PkK2Zc13mqF5iMEQq8CAwEAAaOC
# AV8wggFbMA4GA1UdDwEB/wQEAwIHgDBMBgNVHSAERTBDMEEGCSsGAQQBoDIBHjA0
# MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0
# b3J5LzAJBgNVHRMEAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7
# MDkwN6A1oDOGMWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vZ3MvZ3N0aW1lc3Rh
# bXBpbmdnMi5jcmwwVAYIKwYBBQUHAQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8v
# c2VjdXJlLmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc3RpbWVzdGFtcGluZ2cyLmNy
# dDAdBgNVHQ4EFgQU1KKESjhaGH+6TzBQvZ3VeofWCfcwHwYDVR0jBBgwFoAURtg+
# /9zjvv+D5vSFm7DdatYUqcEwDQYJKoZIhvcNAQEFBQADggEBAIAy3AeNHKCcnTwq
# 6D0hi1mhTX7MRM4Dvn6qvMTme3O7S/GI2pBOdTcoOGO51ysPVKlWznc5lzBzzZvZ
# 2QVFHI2kuANdT9kcLpjg6Yjm7NcFflYqe/cWW6Otj5clEoQbslxjSgrS7xBUR4KE
# NWkonAzkHxQWJPp13HRybk7K42pDr899NkjRvekGkSwvpshx/c+92J0hmPyv294i
# jK+n83fvndyjcEtEGvB4hR7ypYw5tdyIHDftrRT1Bwsmvb5tAl6xuLBYbIU6Dfb/
# WicMxd5T51Q8VkzJTkww9vJc+xqMwoK+rVmR9htNVXvPWwHc/XrTbyNcMkebAfPB
# URRGipswggVMMIIDNKADAgECAhMzAAAANdjVWVsGcUErAAAAAAA1MA0GCSqGSIb3
# DQEBBQUAMH8xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAn
# BgNVBAMTIE1pY3Jvc29mdCBDb2RlIFZlcmlmaWNhdGlvbiBSb290MB4XDTEzMDgx
# NTIwMjYzMFoXDTIzMDgxNTIwMzYzMFowbzELMAkGA1UEBhMCU0UxFDASBgNVBAoT
# C0FkZFRydXN0IEFCMSYwJAYDVQQLEx1BZGRUcnVzdCBFeHRlcm5hbCBUVFAgTmV0
# d29yazEiMCAGA1UEAxMZQWRkVHJ1c3QgRXh0ZXJuYWwgQ0EgUm9vdDCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBALf3GjPm8gAELTngTlvtH7xsD821+iO2
# zt6bETOXpClMfZOfvUq8k+0DGuOPz+VtUFrWlymUWoCwSXrbLpX9uMq/NzgtHj6R
# Qa1wVsfwTz/oMp50ysiQVOnGXw94nZpAPA6sYapeFI+eh6FqUNzXmk6vBbOmcZSc
# cbNQYArHE504B4YCqOmoaSYYkKtMsE8jqzpPhNjfzp/haW+710LXa0Tkx63ubUFf
# clpxCDezeWWkWaCUN/cALw3CknLa0Dhy2xSoRcRdKn23tNbE7qzNE0S3ySvdQwAl
# +mG5aWpYIxG3pzOPVnVZ9c0p10a3CitlttNCbxWyuHv77+ldU9U0WicCAwEAAaOB
# 0DCBzTATBgNVHSUEDDAKBggrBgEFBQcDAzASBgNVHRMBAf8ECDAGAQH/AgECMB0G
# A1UdDgQWBBStvZh6NLQm9/rEJlTvA73gJMtUGjALBgNVHQ8EBAMCAYYwHwYDVR0j
# BBgwFoAUYvsKIVt/Q24R2glUUGv10pZx8Z4wVQYDVR0fBE4wTDBKoEigRoZEaHR0
# cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0
# Q29kZVZlcmlmUm9vdC5jcmwwDQYJKoZIhvcNAQEFBQADggIBADYrovLhMx/kk/fy
# aYXGZA7Jm2Mv5HA3mP2U7HvP+KFCRvntak6NNGk2BVV6HrutjJlClgbpJagmhL7B
# vxapfKpbBLf90cD0Ar4o7fV3x5v+OvbowXvTgqv6FE7PK8/l1bVIQLGjj4OLrSsl
# U6umNM7yQ/dPLOndHk5atrroOxCZJAC8UP149uUjqImUk/e3QTA3Sle35kTZyd+Z
# BapE/HSvgmTMB8sBtgnDLuPoMqe0n0F4x6GENlRi8uwVCsjq0IT48eBr9FYSX5Xg
# /N23dpP+KUol6QQA8bQRDsmEntsXffUepY42KRk6bWxGS9ercCQojQWj2dUk8vig
# 0TyCOdSogg5pOoEJ/Abwx1kzhDaTBkGRIywipacBK1C0KK7bRrBZG4azm4foSU45
# C20U30wDMB4fX3Su9VtZA1PsmBbg0GI1dRtIuH0T5XpIuHdSpAeYJTsGm3pOam9E
# hk8UTyd5Jz1Qc0FMnEE+3SkMc7HH+x92DBdlBOvSUBCSQUns5AZ9NhVEb4m/aX35
# TUDBOpi2oH4x0rWuyvtT1T9Qhs1ekzttXXyaPz/3qSVYhN0RSQCix8ieN913jm1x
# i+BbgTRdVLrM9ZNHiG3n71viKOSAG0DkDyrRfyMVZVqsmZRDP0ZVJtbE+oiV4pGa
# oy0Lhd6sjOD5Z3CfcXkCMfdhoinEMIIFaTCCBFGgAwIBAgIQK1xBJ0ChlqTlel5/
# XafHajANBgkqhkiG9w0BAQsFADB9MQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3Jl
# YXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQKExFDT01P
# RE8gQ0EgTGltaXRlZDEjMCEGA1UEAxMaQ09NT0RPIFJTQSBDb2RlIFNpZ25pbmcg
# Q0EwHhcNMTUxMjE3MDAwMDAwWhcNMTYxMjE2MjM1OTU5WjCB0jELMAkGA1UEBhMC
# VVMxDjAMBgNVBBEMBTk0MzA0MQswCQYDVQQIDAJDQTESMBAGA1UEBwwJUGFsbyBB
# bHRvMRwwGgYDVQQJDBMzMDAwIEhhbm92ZXIgU3RyZWV0MSswKQYDVQQKDCJIZXds
# ZXR0IFBhY2thcmQgRW50ZXJwcmlzZSBDb21wYW55MRowGAYDVQQLDBFIUCBDeWJl
# ciBTZWN1cml0eTErMCkGA1UEAwwiSGV3bGV0dCBQYWNrYXJkIEVudGVycHJpc2Ug
# Q29tcGFueTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMuRzybUEbn9
# Y5tKWlD7eAAWRzoVt4RAGpBHY863qtquCVG0Wtq80GDIIL0WVgnxB+kgLBzeXwU3
# s9YXI37QaBYJt5NDuGEdQ2qpNT64V9ZuKZrhTg64sNx6ONmquRYkXdHp6mVZ3Rvr
# 9OqEg3plx+VMGow5Vfa0zfobabpR3MUgQsQSmB5MP/cmNVVrkzdQmgeYxgXpGxKi
# yYDkTMlLVKaX0NoGnx3HQj+Yw7oILbV9veff4rfjz7BRDEsXa/pjwtOlTgCO6sv8
# J7HMdM7v2+qRNA944cBnxAAR2VlrG50/vkUgwMsPv4H/16pwbkzxvap5OspNhs6o
# 7G/5MF/7DI8CAwEAAaOCAY0wggGJMB8GA1UdIwQYMBaAFCmRYP+KTfrr+aZquM/5
# 5ku9Sc4SMB0GA1UdDgQWBBQ7WJcg/m9cEGQtGzvFwl5wvY5r4zAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzARBglghkgB
# hvhCAQEEBAMCBBAwRgYDVR0gBD8wPTA7BgwrBgEEAbIxAQIBAwIwKzApBggrBgEF
# BQcCARYdaHR0cHM6Ly9zZWN1cmUuY29tb2RvLm5ldC9DUFMwQwYDVR0fBDwwOjA4
# oDagNIYyaHR0cDovL2NybC5jb21vZG9jYS5jb20vQ09NT0RPUlNBQ29kZVNpZ25p
# bmdDQS5jcmwwdAYIKwYBBQUHAQEEaDBmMD4GCCsGAQUFBzAChjJodHRwOi8vY3J0
# LmNvbW9kb2NhLmNvbS9DT01PRE9SU0FDb2RlU2lnbmluZ0NBLmNydDAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBCwUAA4IB
# AQCJSO8rO8/OixqDrdSrsj+AO1UH4zLhQzv/K8OPWSw2+PkyEfvPoe4J6JJ5mNVQ
# 9fWNvFatUv6XcZJ5bv6SmQ0vWbXHNrMvrBtq9hGvJJKFqRhEfz0YM9yTHJUIFMUg
# aAVLRt6/b9k8lJkVPy5IghVGZ5G0AlDpovzZKBxYfJlEEJc6hkjjGBMjkj3ABd21
# jiuAITinnnBsUfUFehZPQSEHI8mlPyB6QboZk7Lz4Yy7emfcfFZB2s7qaWRhZrzK
# kFnzUUZuYr+sEcMPUvORC+qPXSLI9xGH1Y6v88g7DL19bVtODXg3k6BcmxLfPuX+
# CkksoTIYIPX772dsbTfGU+wBMIIFdDCCBFygAwIBAgIQJ2buVutJ846r13Ci/ITe
# IjANBgkqhkiG9w0BAQwFADBvMQswCQYDVQQGEwJTRTEUMBIGA1UEChMLQWRkVHJ1
# c3QgQUIxJjAkBgNVBAsTHUFkZFRydXN0IEV4dGVybmFsIFRUUCBOZXR3b3JrMSIw
# IAYDVQQDExlBZGRUcnVzdCBFeHRlcm5hbCBDQSBSb290MB4XDTAwMDUzMDEwNDgz
# OFoXDTIwMDUzMDEwNDgzOFowgYUxCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVh
# dGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9E
# TyBDQSBMaW1pdGVkMSswKQYDVQQDEyJDT01PRE8gUlNBIENlcnRpZmljYXRpb24g
# QXV0aG9yaXR5MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAkehUktIK
# VrGsDSTdxc9EZ3SZKzejfSNwAHG8U9/E+ioSj0t/EFa9n3Byt2F/yUsPF6c947AE
# Ye7/EZfH9IY+Cvo+XPmT5jR62RRr55yzhaCCenavcZDX7P0N+pxs+t+wgvQUfvm+
# xKYvT3+Zf7X8Z0NyvQwA1onrayzT7Y+YHBSrfuXjbvzYqOSSJNpDa2K4Vf3qwbxs
# tovzDo2a5JtsaZn4eEgwRdWt4Q08RWD8MpZRJ7xnw8outmvqRsfHIKCxH2XeSAi6
# pE6p8oNGN4Tr6MyBSENnTnIqm1y9TBsoilwie7SrmNnu4FGDwwlGTm0+mfqVF9p8
# M1dBPI1R7Qu2XK8sYxrfV8g/vOldxJuvRZnio1oktLqpVj3Pb6r/SVi+8Kj/9Lit
# 6Tf7urj0Czr56ENCHonYhMsT8dm74YlguIwoVqwUHZwK53Hrzw7dPamWoUi9PPev
# tQ0iTMARgexWO/bTouJbt7IEIlKVgJNp6I5MZfGRAy1wdALqi2cVKWlSArvX31Bq
# VUa/oKMoYX9w0MOiqiwhqkfOKJwGRXa/ghgntNWutMtQ5mv0TIZxMOmm3xaG4Nj/
# QN370EKIf6MzOi5cHkERgWPOGHFrK+ymircxXDpqR+DDeVnWIBqv8mqYqnK8V0rS
# S527EPywTEHl7R09XiidnMy/s1Hap0flhFMCAwEAAaOB9DCB8TAfBgNVHSMEGDAW
# gBStvZh6NLQm9/rEJlTvA73gJMtUGjAdBgNVHQ4EFgQUu69+Aj36pvE8hI6t7jiY
# 7NkyMtQwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEQYDVR0gBAow
# CDAGBgRVHSAAMEQGA1UdHwQ9MDswOaA3oDWGM2h0dHA6Ly9jcmwudXNlcnRydXN0
# LmNvbS9BZGRUcnVzdEV4dGVybmFsQ0FSb290LmNybDA1BggrBgEFBQcBAQQpMCcw
# JQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcN
# AQEMBQADggEBAGS/g/FfmoXQzbihKVcN6Fr30ek+8nYEbvFScLsePP9NDXRqzIGC
# JdPDoCpdTPW6i6FtxFQJdcfjJw5dhHk3QBN39bSsHNA7qxcS1u80GH4r6XnTq1dF
# DK8o+tDb5VCViLvfhVdpfZLYUspzgb8c8+a4bmYRBbMelC1/kZWSWfFMzqORcUx8
# Rww7Cxn2obFshj5cqsQugsv5B5a6SE2Q8pTIqXOi6wZ7I53eovNNVZ96YUWYGGjH
# XkBrI/V5eu+MtWuLt29G9HvxPUsE2JOAWVrgQSQdso8VYFhH2+9uRv0V9dlfmrPb
# 2LjkQLPNlzmuhbsdjrzch5vRpu/xO28QOG8wggXgMIIDyKADAgECAhAufIfMDpNK
# Uv6U/Ry3zTSvMA0GCSqGSIb3DQEBDAUAMIGFMQswCQYDVQQGEwJHQjEbMBkGA1UE
# CBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQK
# ExFDT01PRE8gQ0EgTGltaXRlZDErMCkGA1UEAxMiQ09NT0RPIFJTQSBDZXJ0aWZp
# Y2F0aW9uIEF1dGhvcml0eTAeFw0xMzA1MDkwMDAwMDBaFw0yODA1MDgyMzU5NTla
# MH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAO
# BgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBMaW1pdGVkMSMwIQYD
# VQQDExpDT01PRE8gUlNBIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAKaYkGN3kTR/itHd6WcxEevMHv0xHbO5Ylc/k7xb458e
# JDIRJ2u8UZGnz56eJbNfgagYDx0eIDAO+2F7hgmz4/2iaJ0cLJ2/cuPkdaDlNSOO
# yYruGgxkx9hCoXu1UgNLOrCOI0tLY+AilDd71XmQChQYUSzm/sES8Bw/YWEKjKLc
# 9sMwqs0oGHVIwXlaCM27jFWM99R2kDozRlBzmFz0hUprD4DdXta9/akvwCX1+XjX
# jV8QwkRVPJA8MUbLcK4HqQrjr8EBb5AaI+JfONvGCF1Hs4NB8C4ANxS5Eqp5klLN
# hw972GIppH4wvRu1jHK0SPLj6CH5XkxieYsCBp9/1QsCAwEAAaOCAVEwggFNMB8G
# A1UdIwQYMBaAFLuvfgI9+qbxPISOre44mOzZMjLUMB0GA1UdDgQWBBQpkWD/ik36
# 6/mmarjP+eZLvUnOEjAOBgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIB
# ADATBgNVHSUEDDAKBggrBgEFBQcDAzARBgNVHSAECjAIMAYGBFUdIAAwTAYDVR0f
# BEUwQzBBoD+gPYY7aHR0cDovL2NybC5jb21vZG9jYS5jb20vQ09NT0RPUlNBQ2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eS5jcmwwcQYIKwYBBQUHAQEEZTBjMDsGCCsGAQUF
# BzAChi9odHRwOi8vY3J0LmNvbW9kb2NhLmNvbS9DT01PRE9SU0FBZGRUcnVzdENB
# LmNydDAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqG
# SIb3DQEBDAUAA4ICAQACPwI5w+74yjuJ3gxtTbHxTpJPr8I4LATMxWMRqwljr6ui
# 1wI/zG8Zwz3WGgiU/yXYqYinKxAa4JuxByIaURw61OHpCb/mJHSvHnsWMW4j71RR
# LVIC4nUIBUzxt1HhUQDGh/Zs7hBEdldq8d9YayGqSdR8N069/7Z1VEAYNldnEc1P
# AuT+89r8dRfb7Lf3ZQkjSR9DV4PqfiB3YchN8rtlTaj3hUUHr3ppJ2WQKUCL33s6
# UTmMqB9wea1tQiCizwxsA4xMzXMHlOdajjoEuqKhfB/LYzoVp9QVG6dSRzKp9L9k
# R9GqH1NOMjBzwm+3eIKdXP9Gu2siHYgL+BuqNKb8jPXdf2WMjDFXMdA27Eehz8uL
# qO8cGFjFBnfKS5tRr0wISnqP4qNS4o6OzCbkstjlOMKo7caBnDVrqVhhSgqXtEtC
# tlWdvpnncG1Z+G0qDH8ZYF8MmohsMKxSCZAWG/8rndvQIMqJ6ih+Mo4Z33tIMx7X
# ZfiuyfiDFJN2fWTQjs6+NX3/cjFNn569HmwvqI8MBlD7jCezdsn05tfDNOKMhyGG
# Yf6/VXThIXcDCmhsu+TJqebPWSXrfOxFDnlmaOgizbjvmIVNlhE8CYrQf7woKBP7
# aspUjZJczcJlmAaezkhb1LU3k0ZBfAfdz/pD77pnYf99SeC7MH1cgOPmFjlLpzGC
# BOEwggTdAgEBMIGRMH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1h
# bmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGjAYBgNVBAoTEUNPTU9ETyBDQSBM
# aW1pdGVkMSMwIQYDVQQDExpDT01PRE8gUlNBIENvZGUgU2lnbmluZyBDQQIQK1xB
# J0ChlqTlel5/XafHajANBglghkgBZQMEAgEFAKB8MBAGCisGAQQBgjcCAQwxAjAA
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDzF6Hl6rCu4sBDJ2zSXUZgLIkQSRv0
# 6gTlBQJlmQqBhDANBgkqhkiG9w0BAQEFAASCAQClvhKLj6JQn515HdFuIg6TwLDm
# BrYJMYStQMghL5c1CvX+JeXUAt0BQz7nv5eKPj5h6nBJyXn+7ZPWbfaIdZyws7mI
# bWdOIEJWmxZkXkcykZBCSn2Qk1czg4aOIO22mScOgFuyAI0Lu2cw63u7HIP9ThMQ
# x1yrsuJAjEROtr7BRLjdh15gNGSlEQDIOJYRFDeF+dqNOCpH8RGAF4frV0dEKb+B
# mN0cxLFsjQbcxthq+vbUzApwiTckhWYpsbipWQdi7kYWIypo4XkN5JEorqtWlZbS
# skBlApscADUXrBSlSDVDmBkv5GbxUDxOpmEQVpyfP+NLcliT12BiH3e7Rtg1oYIC
# ojCCAp4GCSqGSIb3DQEJBjGCAo8wggKLAgEBMGgwUjELMAkGA1UEBhMCQkUxGTAX
# BgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGlt
# ZXN0YW1waW5nIENBIC0gRzICEhEhBqCB0z/YeuWCTMFrUglOAzAJBgUrDgMCGgUA
# oIH9MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2
# MDUxMzE4Mzk0M1owIwYJKoZIhvcNAQkEMRYEFAeQVGznZtcMA2KoFE1+3YM4RxTt
# MIGdBgsqhkiG9w0BCRACDDGBjTCBijCBhzCBhAQUs2MItNTN7U/PvWa5Vfrjv7Es
# KeYwbDBWpFQwUjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYt
# c2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzICEhEh
# BqCB0z/YeuWCTMFrUglOAzANBgkqhkiG9w0BAQEFAASCAQB9Jnrd8eWO3W+IaohK
# VX23ip7DI//fKiEtkfa542/gJKBtDaRkFt1ZhM8bjRYh04pG8YFbqMif9DGtyX6K
# 9QaRZ/Ko+t9UziN1yq27jFuc2n8R6jxJ/jWydF4ze5k3WnbuQN4lqkYbEbemEPuZ
# XX11ncuEsP6GP81AMcwpHzXYfOn6+fTIOyKa6lYiPgUtVvNNUvQ2QPoHP0DRSbrO
# G/xSD6hiTaVAyzpCHGxI2N8V5t9C8Yojq+Wq5Yuiifv8W6OyCE7s7ZsfUfJ1+WQ9
# M1a0bLae/Y3hVbe1xZcZwSBQcZfY85yxM32GsLj9wn1MMgHdI1lXtjnowzvWmB3K
# N4jK
# SIG # End signature block
