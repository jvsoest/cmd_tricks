# See instructions at https://www.thewindowsclub.com/change-registry-using-windows-powershell
# NOTE: THIS SCRIPT NEEDS TO BE EXECUTED WITH UAC/ADMIN PRIVILEGES

# This override makes it possible to use the exchange cached mode in Outlook
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\outlook\cached mode" -Name enable -Value 1 -Force

# This override makes it possible to use VBA macros in Outlook. Usual GPO value is 4 (allow none) where
#   this override value (2) indicates that scripts are allowed but all scripts are notified beforehand.
#   Usually this is setup in Outlook -> File -> Options -> Trust Center
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\outlook\security" -Name level -Value 2 -Force

# This override makes it possible to use VBA macros in Word. Usual GPO value is 4 (blocked). This causes Zotero to not work properly.
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\word\security" -Name vbawarnings 1 -Force