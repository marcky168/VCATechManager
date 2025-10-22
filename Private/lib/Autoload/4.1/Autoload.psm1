#Requires -Version 2.0
## Automatically load functions from scripts on-demand, instead of having to dot-source them ahead of time, or reparse them from the script every time.
## Provides significant memory benefits over pre-loading all your functions, and significant performance benefits over using plain scripts.  Can also *inject* functions into Modules so they inherit the module scope instead of the current local scope.
## Please see the use example in the script below

## Version History
## v 1.2  - 2011.05.02
##        - Exposed the LoadNow alias and the Resolve-Autoloaded function
## v 1.1  - 2011.02.09
##          Added support for autoloading scripts (files that don't have a "function" in them)
## v 1.0  - 2010.10.20
##          Officially out of beta -- this is working for me without problems on a daily basis.
##          Renamed functions to respect the Verb-Noun expectations, and added Export-ModuleMember
## beta 8 - 2010.09.20
##          Finally fixed the problem with aliases that point at Invoke-Autoloaded functions!
## beta 7 - 2010.06.03
##          Added some help, and a function to force loading "now"
##          Added a function to load AND show the help...
## beta 6 - 2010.05.18
##          Fixed a bug in output when multiple outputs happen in the END block
## beta 5 - 2010.05.10
##          Fixed non-pipeline use using $MyInvocation.ExpectingInput
## beta 4 - 2010.05.10
##          I made a few tweaks and bug fixes while testing it's use with PowerBoots.
## beta 3 - 2010.05.10
##          fix for signed scripts (strip signature)
## beta 2 - 2010.05.09
##          implement module support
## beta 1 - 2010.04.14
##          Initial Release


## To use:
## 1) Create a function. To be 100% compatible, it should specify pipeline arguments
## For example:
<#
function Skip-Object {
   param( 
      [int]$First = 0, [int]$Last = 0, [int]$Every = 0, [int]$UpTo = 0,  
      [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
      $InputObject
   )
   begin {
      if($Last) {
         $Queue = new-object System.Collections.Queue $Last
      }
      $e = $every; $UpTo++; $u = 0
   }
   process {
      $InputObject | where { --$First -lt 0 } | 
      foreach {
         if($Last) {
            $Queue.EnQueue($_)
            if($Queue.Count -gt $Last) { $Queue.DeQueue() }
         } else { $_ }
      } |
      foreach { 
         if(!$UpTo) { $_ } elseif( --$u -le 0) {  $_; $U = $UpTo }
      } |
      foreach { 
         if($every -and (--$e -le 0)) {  $e = $every  } else { $_ } 
      }
   }
}
#>

## 2) Put the function into a script (for our example: C:\Users\${Env:UserName}\Documents\WindowsPowerShell\Scripts\SkipObject.ps1 )
## 3) Import the Autoload Module
## 5) Run this command (or add it to your profile):
<#
New-Autoload C:\Users\${Env:UserName}\Documents\WindowsPowerShell\Scripts\SkipObject.ps1 Skip-Object
#>

## This tells us that you want to have that function loaded for you out of the script file if you ever try to use it.
## Now, you can just use the function:
## 1..10 | Skip-Object -first 2 -upto 2

function Invoke-Autoloaded {
   #.Synopsis
   #	This function was autoloaded, but it has not been parsed yet.
   #  Use Get-AutoloadHelp to force parsing and get the correct help (or just invoke the function once).
   #.Description
   #   You are seeing this help because the command you typed was imported via the New-Autoload command from the Autoload module.  The script file containing the function has not been loaded nor parsed yet. In order to see the correct help for your function we will need to parse the full script file, to force that at this time you may use the Get-AutoloadHelp function.
   #
   #   For example, if your command was Get-PerformanceHistory, you can force loading the help for it by running the command: Get-AutoloadHelp Get-PerformanceHistory
   [CmdletBinding()]Param()
   DYNAMICPARAM {
      $CommandName = $MyInvocation.InvocationName
	   return Resolve-Autoloaded $CommandName
   }#DynamicParam

   begin {
      Write-Verbose "Command: $CommandName"
      if(!$Script:AutoloadHash[$CommandName]) {
         do {
            $Alias = $CommandName
            $CommandName = Get-Alias $CommandName -ErrorAction SilentlyContinue | Select -Expand Definition
            Write-Verbose "Invoke-Autoloaded Begin: $Alias -> $CommandName"
         } while(!$Script:AutoloadHash[$CommandName] -and (Get-Alias $CommandName -ErrorAction SilentlyContinue))
      } else {
         Write-Verbose "CommandHash: $($Script:AutoloadHash[$CommandName])"
      }
      if(!$Script:AutoloadHash[$CommandName]) { throw "Unable to determine command!" }

      $ScriptName, $ModuleName, $FunctionName = $Script:AutoloadHash[$CommandName]
      Write-Verbose "Invoke-Autoloaded Begin: $Alias -> $CommandName -> $FunctionName"
      
      
      #Write-Host "Parameters: $($PSBoundParameters | ft | out-string)" -Fore Magenta
   
      $global:command = $ExecutionContext.InvokeCommand.GetCommand( $FunctionName, [System.Management.Automation.CommandTypes]::Function )
      Write-Verbose "Autoloaded Command: $($Command|Out-String)"
      $scriptCmd = {& $command @PSBoundParameters | Write-Output }
      $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
      $steppablePipeline.Begin($myInvocation.ExpectingInput)
   }
   process
   {
      Write-Verbose "Invoke-Autoloaded Process: $CommandName ($_)"
      try {
         if($_) {
            $steppablePipeline.Process($_)
         } else {
            $steppablePipeline.Process()
         }
      } catch {
         throw
      }
   }

   end
   {
      try {
         $steppablePipeline.End()
      } catch {
         throw
      }
      Write-Verbose "Invoke-Autoloaded End: $CommandName"
   }
}#Invoke-Autoloaded


function Get-AutoloadHelp {
	[CmdletBinding()]
	Param([Parameter(Mandatory=$true)][String]$CommandName)
	$null = Resolve-Autoloaded $CommandName
	Get-Help $CommandName
}

function Resolve-Autoloaded {
   [CmdletBinding()]
   param(
      [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
      [Alias("Name")]
      [String]$CommandName
   )
   $OriginalCommandName = "($CommandName)"
   Write-Verbose "Command: $CommandName"
   if(!$Script:AutoloadHash[$CommandName]) {
      do {
         $Alias = $CommandName
         $CommandName = Get-Alias $CommandName -ErrorAction SilentlyContinue | Select -Expand Definition
         $OriginalCommandName += "($CommandName)"
         Write-Verbose "Resolve-Autoloaded Begin: $Alias -> $CommandName"
      } while(!$Script:AutoloadHash[$CommandName] -and (Get-Alias $CommandName -ErrorAction SilentlyContinue))
   } else {
      Write-Verbose "CommandHash: $($Script:AutoloadHash[$CommandName])"
   }
   if(!$Script:AutoloadHash[$CommandName]) { throw "Unable to determine command $OriginalCommandName!" }
   
   Write-Verbose "Resolve-Autoloaded DynamicParam: $CommandName from $($Script:AutoloadHash[$CommandName])"
   $ScriptName, $ModuleName, $FunctionName = $Script:AutoloadHash[$CommandName]
   Write-Verbose "Autoloading:`nScriptName: $ScriptName `nModuleName: $ModuleName `nFunctionName: $FunctionName"
   
   if(!$ScriptName){ $ScriptName = $CommandName }
   if(!$FunctionName){ $FunctionName = $CommandName }
   if($ModuleName) {
      $Module = Get-Module $ModuleName
   } else { $Module = $null }
   
   
   ## Determine the command name based on the alias used to invoke us
   ## Store the parameter set for use in the function later...
   $paramDictionary = new-object System.Management.Automation.RuntimeDefinedParameterDictionary
   
   #$externalScript = $ExecutionContext.InvokeCommand.GetCommand( $CommandName, [System.Management.Automation.CommandTypes]::ExternalScript )
   $externalScript = Get-Command $ScriptName -Type ExternalScript | Select -First 1
   Write-Verbose "Processing Script: $($externalScript |Out-String)"
   $parserrors = $null
   $prev = $null
   $script = $externalScript.ScriptContents
   [System.Management.Automation.PSToken[]]$tokens = [PSParser]::Tokenize( $script, [ref]$parserrors )
   [Array]::Reverse($tokens)
   
   $function = $false
   foreach($token in $tokens) {
      if($prev -and $token.Content -eq "# SIG # Begin signature block") {
         $script = $script.SubString(0, $token.Start )
      }
      if($prev -and $token.Type -eq "Keyword" -and $token.Content -ieq "function" -and $prev.Content -ieq $FunctionName ) {
         $script = $script.Insert( $prev.Start, "global:" )
         $function = $true
         Write-Verbose "Globalized: $($script[(($prev.Start+7)..($prev.Start + 7 +$prev.Content.Length))] -join '')"
      }
      $prev = $token
   }
   
   if(!$function) {
      $script = "function global:$Functionname { $script }"
   }
   
   if($Module) {
      $script = Invoke-Expression "{ $Script }"
      Write-Verbose "Importing Function into $($Module) module."
      &$Module $Script | Out-Null
      $command = Get-Command $FunctionName -Type Function
      Write-Verbose "Loaded Module Function: $($command | ft CommandType, Name, ModuleName, Visibility|Out-String)"
   } else {
      Write-Verbose "Importing Function without module."
      Invoke-Expression $script | out-null
      $command = Get-Command $FunctionName -Type Function
      Write-Verbose "Loaded Local Function: $($command | ft CommandType, Name, ModuleName, Visibility|Out-String)"
   }
   if(!$command) {
      throw "Something went wrong autoloading the $($FunctionName) function. Function definition doesn't exist in script: $($externalScript.Path)"
   }
   
   if($CommandName -eq $FunctionName) {
      Remove-Item Alias::$($CommandName)
      Write-Verbose "Defined the function $($FunctionName) and removed the alias $($CommandName)"
   } else {
      Set-Alias $CommandName $FunctionName -Scope Global
      Write-Verbose "Defined the function $($FunctionName) and redefined the alias $($CommandName)"
   }
   foreach( $pkv in $command.Parameters.GetEnumerator() ){
      $parameter = $pkv.Value
      if( $parameter.Aliases -match "vb|db|ea|wa|ev|wv|ov|ob" ) { continue } 
      $param = new-object System.Management.Automation.RuntimeDefinedParameter( $parameter.Name, $parameter.ParameterType, $parameter.Attributes)
      $paramdictionary.Add($pkv.Key, $param)
   } 
   return $paramdictionary
}

function New-Autoload {
   [CmdletBinding()]
   param(
      [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
      [string[]]$Name,

      [Parameter(Position=1,Mandatory=$False,ValueFromPipelineByPropertyName=$true)]
      [Alias("BaseName")]
      $Alias = $Name,

      [Parameter(Position=2,Mandatory=$False,ValueFromPipelineByPropertyName=$true)]
      $Function = $Alias,

      [Parameter(Position=3,Mandatory=$false)]
      [String]$Module,

      [Parameter(Mandatory=$false)]
      [String]$Scope = '2'
   )
   begin {
      $xlr8r = [psobject].assembly.gettype("System.Management.Automation.TypeAccelerators")
      if(!$xlr8r::Get["PSParser"]) {
         if($xlr8r::AddReplace) { 
            $xlr8r::AddReplace( "PSParser", "System.Management.Automation.PSParser, System.Management.Automation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" )
         } else {
            $null = $xlr8r::Remove( "PSParser" )
            $xlr8r::Add( "PSParser", "System.Management.Automation.PSParser, System.Management.Automation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" )
         }

      
      }   
   }
   process {
      for($i=0;$i -lt $Name.Count;$i++){
         if($Alias -is [Scriptblock]) {
            $a = $Name[$i] | &$Alias
         } elseif($Alias -is [Array]) {
            $a = $Alias[$i]
         } else {
            $a = $Alias
         }
         
         if($Function -is [Scriptblock]) {
            $f = $Name[$i] | &$Function
         } elseif($Function -is [Array]) {
            $f = $Function[$i]
         } else {
            $f = $Function
         }
         
         Write-Verbose "Set-Alias $Module\$a Invoke-Autoloaded -Scope $Scope"
         Set-Alias $a Invoke-Autoloaded -Scope $Scope
         $Script:AutoloadHash[$a] = $Name[$i],$Module,$f
         Write-Verbose "$($Script:AutoloadHash.Count)  `$AutoloadHash[$a] = $($Script:AutoloadHash[$a])"
      }
   }
}

Set-Alias Autoload New-Autoload
Set-Alias LoadNow  Resolve-Autoloaded

New-Variable -Name AutoloadHash -Value @{} -Scope Script -Description "The Autoload alias-to-script cache" -Option ReadOnly -ErrorAction SilentlyContinue

Export-ModuleMember -Function New-Autoload, Invoke-Autoloaded, Get-AutoloadHelp, Resolve-Autoloaded -Alias *


# SIG # Begin signature block
# MIIXxAYJKoZIhvcNAQcCoIIXtTCCF7ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUa3zb0fCmJxC+Y1wOHWrCSihH
# m9mgghL3MIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggUmMIIEDqADAgECAhACXbrxBhFj1/jVxh2rtd9BMA0GCSqGSIb3DQEBCwUAMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0EwHhcNMTUwNTA0MDAwMDAwWhcNMTYwNTExMTIwMDAw
# WjBtMQswCQYDVQQGEwJVUzERMA8GA1UECBMITmV3IFlvcmsxFzAVBgNVBAcTDldl
# c3QgSGVucmlldHRhMRgwFgYDVQQKEw9Kb2VsIEguIEJlbm5ldHQxGDAWBgNVBAMT
# D0pvZWwgSC4gQmVubmV0dDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# AJfRKhfiDjMovUELYgagznWf+HFcDENk118Y/K6UkQDwKmVyVOvDyaVefjSmZZcV
# NZqqYpm9d/Iajf2dauyC3pg3oay8KfXAADLHgbmbvYDc5zGuUNsTzMUOKlp9h13c
# qsg898JwpRpI659xCQgJjZ6V83QJh+wnHvjA9ojjA4xkbwhGp4Eit6B/uGthEA11
# IHcFcXeNI3fIkbwWiAw7ZoFtSLm688NFhxwm+JH3Xwj0HxuezsmU0Yc/po31CoST
# nGPVN8wppHYZ0GfPwuNK4TwaI0FEXxwdwB+mEduxa5e4zB8DyUZByFW338XkGfc1
# qcJJ+WTyNKFN7saevhwp02cCAwEAAaOCAbswggG3MB8GA1UdIwQYMBaAFFrEuXsq
# CqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBQV0aryV1RTeVOG+wlr2Z2bOVFAbTAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1
# oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1n
# MS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMEIGA1UdIAQ7MDkwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwgYQGCCsGAQUFBwEBBHgw
# djAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUF
# BzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNz
# dXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0B
# AQsFAAOCAQEAIi5p+6eRu6bMOSwJt9HSBkGbaPZlqKkMd4e6AyKIqCRabyjLISwd
# i32p8AT7r2oOubFy+R1LmbBMaPXORLLO9N88qxmJfwFSd+ZzfALevANdbGNp9+6A
# khe3PiR0+eL8ZM5gPJv26OvpYaRebJTfU++T1sS5dYaPAztMNsDzY3krc92O27AS
# WjTjWeILSryqRHXyj8KQbYyWpnG2gWRibjXi5ofL+BHyJQRET5pZbERvl2l9Bo4Z
# st8CM9EQDrdG2vhELNiA6jwenxNPOa6tPkgf8cH8qpGRBVr9yuTMSHS1p9Rc+ybx
# FSKiZkOw8iCR6ZQIeKkSVdwFf8V+HHPrETCCBTAwggQYoAMCAQICEAQJGBtf1btm
# dVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UE
# AxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEyMDAwMFoX
# DTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNl
# cnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE620T1f+Wo
# ndsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG+yknx9N7
# I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwgla4cSocI
# 3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q16XTmK/5s
# y350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTFjg/Bougs
# UfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOCAc0wggHJ
# MBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4
# MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZIAYb9bAAC
# BDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAoG
# CGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOCAQEAPuwN
# WiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3sOH20ZJ1
# D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsHDpEV+7qv
# tVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlxsQyPu6j4
# xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy62ZD2rOw
# jNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQYhS6Skepo
# bEQysmah5xikmmRR7zGCBDcwggQzAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAJduvEGEWPX+NXGHau130EwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCEbH+BHXSgBZdxL7t1c
# NuIKNep5MA0GCSqGSIb3DQEBAQUABIIBAHsAR/p7OEbbhc+9znZDtc9IpVteXDqe
# DW6gBkQHeqcwyCkU9+TN8yKPGG7pXKNpEWL1v9nSyA3jB5d3AQ8AzMyysaA8KiAv
# C7+973unxb3YPcKLIA/2vMUqweV5n5081feyPJ1kiJ5rDA3ab2H+H3+0NgtZ2ipy
# qb5XVHyhsnbbIm9B/xpoHUyO0O9qfjjmVtaaMSpqr9cCviZpfi5Rcz3GJXL/8nLo
# gPgYuJpDP9G4c6QAWjZygvEQq7gUnRUcg2ZBPT46r2UpQt/5XBAi4Dyb/3hZn+ZL
# lUuRD867D/u4TeUS3MXaXqaNDHl6/jivaNhndgTIQa3RaIRG++RVZtmhggILMIIC
# BwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFudGVjIFRpbWUg
# U3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0OMj+vzVuBNhqmBsaUDAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTUwNTA1MDE1NDM5WjAjBgkqhkiG9w0BCQQxFgQUpPFrkfUh+Xb68ucWzfRT
# qLH5DWcwDQYJKoZIhvcNAQEBBQAEggEACw368c1oSdrCCZdrxJOVYT6fEw5YIkAk
# xAwOSYu3ovX3SMsScidmyC9rvY3Iotooez5VXPGUClNdFoi4QNvXC2C2O64zTn4O
# 5UYcmRfmJrfI5hen+rF2sUHZBPasf7LqpAamBlqVsrCkBeOx2P0OPOgp3ryB8VBX
# E4vHEeAOTAioexnBXjL2LULk1ctHZpMoElLrbtZKXcan7L9F1gFVYtobXz5PK1jE
# zbGJCg07qfFBGFJupK4EvlLQNdFTYkUsL1ZucefAcYMOyTgTUE8nR7621XfPPzCu
# CvqeaVQx2ZFqT8cH8WeQm0pjkDZ4MyFpJNong7VdIkVj1AqHD83UCA==
# SIG # End signature block
