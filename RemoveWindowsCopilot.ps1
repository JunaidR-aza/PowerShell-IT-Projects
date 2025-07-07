<#
.SYNOPSIS
    Removes Windows Copilot for all users

.DESCRIPTION
    This script:
    1. Removes Copilot for all existing users
    2. Removes the provisioned package to prevent new user installations
    3. Requires administrative privileges
#>

# Set execution policy for current process only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force

Write-Host "Starting system-wide Copilot removal..." -ForegroundColor Cyan

try {
    # Remove Copilot for all users
    $packages = Get-AppxPackage -Name "*copilot*" -AllUsers
    if ($packages) {
        Write-Host "Removing Copilot for all users..." -ForegroundColor Yellow
        $packages | Remove-AppxPackage -AllUsers -ErrorAction Stop
        Write-Host "Successfully removed Copilot for all users" -ForegroundColor Green
    }
    else {
        Write-Host "No Copilot installations found for any users" -ForegroundColor Yellow
    }

    # Remove provisioned package (prevents new user installations)
    $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like "*copilot*" }
    if ($provisioned) {
        Write-Host "Removing system provisioned Copilot package..." -ForegroundColor Yellow
        $provisioned | Remove-AppxProvisionedPackage -Online -ErrorAction Stop
        Write-Host "Successfully removed provisioned package" -ForegroundColor Green
    }
    else {
        Write-Host "No provisioned Copilot package found" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error occurred: $_" -ForegroundColor Red
    exit 1
}

Write-Host "`nCopilot removal completed successfully!" -ForegroundColor Green
Write-Host "Changes will take effect immediately for all users." -ForegroundColor Cyan

# Reset execution policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Restricted -Force