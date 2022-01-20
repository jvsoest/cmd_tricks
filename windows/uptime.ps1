$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
$CurrentDate = Get-Date
$uptime = $CurrentDate - $bootuptime
echo "up $($uptime.Days) days, $($uptime.Hours):$($uptime.Minutes)"