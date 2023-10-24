# Uninstall HP Wolf AntiVirus
$hpAntiVirus = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*HP Wolf AntiVirus*" }
if ($hpAntiVirus) {
    $hpAntiVirus.Uninstall()
}

# Uninstall Office 365 (assuming it's a Click-to-Run installation)
$office365 = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft 365%'" | Select-Object -First 1
if ($office365) {
    $office365.Uninstall()
}

# Uninstall Office 2021 (assuming it's a Click-to-Run installation)
$office2021 = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE (Name LIKE 'Microsoft Office%') AND (Name NOT LIKE 'Microsoft 365%')" | Select-Object -First 1
if ($office2021) {
    $office2021.Uninstall()
}

# Open a Chrome link
Start-Process "chrome.exe" "https://dropbox.com/s/q2en360f3gxu2zg/Office2021_ProPlus_English.exe?dl=1"

# Run a specified EXE file in the Downloads folder
$exePath = "$env:LocalAdmin\Downloads\Office2021_ProPlus_English.exe"
if (Test-Path -Path $exePath) {
    Start-Process $exePath
}

# Connect to all printers on the Wi-Fi connection
Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Network = 'True'" | ForEach-Object {
    Invoke-Expression "RUNDLL32 PRINTUI.DLL,PrintUIEntry /in /n `"$($_.Name)`""
}

# Prompt the user for Office 2021 activation key
$activationKey = Read-Host "Enter your Office 2021 activation key"

# Path to the Office 2021 activation script
$activationScriptPath = "C:\Path\To\ActivateOffice2021.vbs"  # Replace with the actual path

# Create the activation script
$activationScript = @"
Set objService = GetObject ("winmgmts:\\\\.\\root\\cimv2")
Set objSoftware = objService.Get ("Win32_Product").SpawnInstance_()
objSoftware.ProductCode = "Your_Product_Code"  # Replace with your Office 2021 product code
objSoftware.InstallLocation = "C:\Program Files\Microsoft Office\Office16"  # Replace with the correct path
objSoftware.InstallSource = "C:\Path\To\Office2021\Setup\Files"  # Replace with the correct path
objSoftware.User = "$env:USERNAME"
objSoftware.Password = "$activationKey"
objSoftware.LicenceKey = "$activationKey"
objSoftware.Put_
"@

# Save the activation script to a VBS file
$activationScript | Out-File -FilePath $activationScriptPath -Encoding ASCII

# Run the activation script with cscript
Start-Process -FilePath "cscript.exe" -ArgumentList $activationScriptPath -Wait

# Remove the activation script
Remove-Item -Path $activationScriptPath -Force

Write-Host "Office 2021 has been activated with the provided key."