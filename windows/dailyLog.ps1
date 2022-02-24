$dateYmd = $(Get-Date).tostring("yyyy_MM_dd")
$fileName = "$HOME\OneDrive\Grip\DailyLogs\$dateYmd.md"
$dateFormatted = $(Get-Date).tostring("dd-MM-yyyy")

$fileExists = Test-Path -Path $fileName -PathType Leaf
if ( !$fileExists ) {
    echo "# $dateFormatted" >> $fileName
    echo "" >> $fileName
}

powershell -windowstyle hidden "code -n $HOME\OneDrive\Grip\; code $fileName"