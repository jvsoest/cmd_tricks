# See instructions at https://www.thewindowsclub.com/change-registry-using-windows-powershell
# NOTE: THIS SCRIPT NEEDS TO BE EXECUTED WITH UAC/ADMIN PRIVILEGES
Set-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\office\16.0\outlook\cached mode" -Name enable -Value 1 -Force