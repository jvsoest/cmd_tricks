$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
$CurrentDate = Get-Date
$uptime = $CurrentDate - $bootuptime
$minutes = $uptime.Minutes
if ($minutes -lt 10) {
    $minutes = "0$($minutes)"
}
echo "up $($uptime.Days) days, $($uptime.Hours):$($minutes)"