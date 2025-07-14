# Set execution policy to unrestricted for the current session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force

# Define the package path
$packagePath = "C:\Temp\Copilot\Microsoft.MicrosoftOfficeHub_19.2507.39061.0_neutral_~_8wekyb3d8bbwe.Msixbundle"

# Run DISM to add the provisioned appx package silently
dism /Online /Add-ProvisionedAppxPackage /PackagePath:"$packagePath" /SkipLicense
