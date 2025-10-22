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
