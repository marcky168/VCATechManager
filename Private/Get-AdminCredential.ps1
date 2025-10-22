# New: Credential Manager
function Get-AdminCredential {
    $cred = Get-StoredCredential -Target "vcaadmin"
    if (-not $cred) {
        $cred = Get-Credential -Message "Enter admin credentials for AD/DHCP"
        New-StoredCredential -Target "vcaadmin" -Credentials $cred -Persist Enterprise
    }
    return $cred
}