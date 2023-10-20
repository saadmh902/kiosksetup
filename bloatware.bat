
msiexec.exe /X {uninstall string} /qn

::uninstalling HP WOLF
::uninstalling office
@echo off

taskkill /f /im HP.ClientSecurityManager.EXE
wmic product where name="HP Wolf Security" call uninstall


::using uninstaller
msiexec.exe /X {uninstall string} /qn
::wipe files
REM Remove Office program files
rd /s /q "C:\Program Files\HP"



::uninstalling office
@echo off
REM Stop all Office applications before proceeding
taskkill /f /im WINWORD.EXE
taskkill /f /im EXCEL.EXE
taskkill /f /im POWERPNT.EXE
taskkill /f /im OUTLOOK.EXE
taskkill /f /im ONEDRIVE.EXE

REM Uninstall Office 365
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /uninstall

REM Uninstall Office 2019
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /uninstall "AllProducts"

REM Uninstall Office 2016
"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /uninstall "ProPlusRetail" /dll OSETUP.DLL

REM Remove Office registry keys
reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office /f
reg delete HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office /f
reg delete HKEY_CLASSES_ROOT\Installer\Products\00002109000000000000000000F01FEC /f

REM Remove Office program files
rd /s /q "C:\Program Files\Microsoft Office"
rd /s /q "C:\Program Files (x86)\Microsoft Office"
rd /s /q "C:\Program Files\Common Files\Microsoft Shared\ClickToRun"
rd /s /q "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun"

echo "Microsoft Office has been uninstalled successfully."

::Add Printer:

rundll32 printui.dll,PrintUIEntry /in /n "\\servername\printer name

::uninstall hpwolf

msiexec.exe /X {uninstall string} /qn


start chrome.exe http://dropbox.com/s/q2en360f3gxu2zg/Office2021_ProPlus_English.exe?dl=1

echo "press enter when install is done"
pause
start "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk"
echo "Enter activation key"
