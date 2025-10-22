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
