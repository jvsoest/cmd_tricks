# As Win11 has some issues with explorer.exe and WinUI, sometimes it can help to give a restart to the process
Stop-Process $(Get-Process -name explorer).Id