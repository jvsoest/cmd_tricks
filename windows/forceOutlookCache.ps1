# See instructions at https://www.thewindowsclub.com/change-registry-using-windows-powershell
# NOTE: THIS SCRIPT NEEDS TO BE EXECUTED WITH UAC/ADMIN PRIVILEGES

# This override makes it possible to use the exchange cached mode in Outlook
# Note: you might need to add the folder "cached mode" and the DWORD item "enable" manually
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\outlook\cached mode" -Name enable -Value 1 -Force

# This override makes it possible to use VBA macros in Outlook. Usual GPO value is 4 (allow none) where
#   this override value (2) indicates that scripts are allowed but all scripts are notified beforehand.
#   Usually this is setup in Outlook -> File -> Options -> Trust Center
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\outlook\security" -Name level -Value 2 -Force

# This override makes it possible to use VBA macros in Word. Usual GPO value is 4 (blocked). This causes Zotero to not work properly.
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\word\security" -Name vbawarnings 1 -Force

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Appx" -Name AllowDeploymentInSpecialProfiles 1 -Force

# Remove group policy that forces the old windows context menu
reg.exe delete "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" /f