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