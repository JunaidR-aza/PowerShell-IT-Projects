# Set execution policy
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force

# Define paths
$networkPath = "\\veris.com.au\data\ICT\02\SYDNEY\IT_Store\Software\Microsoft Office Apps\Copilot\Microsoft.MicrosoftOfficeHub_19.2507.39061.0_neutral_~_8wekyb3d8bbwe.Msixbundle"
$localPath = "$env:TEMP\Microsoft.MicrosoftOfficeHub_19.2507.39061.0_neutral_~_8wekyb3d8bbwe.Msixbundle"

# Copy file locally
Copy-Item -Path $networkPath -Destination $localPath -Force

# Run DISM with local path
dism /Online /Add-ProvisionedAppxPackage /PackagePath:"$localPath" /SkipLicense

# Clean up (optional)
Remove-Item -Path $localPath -Force