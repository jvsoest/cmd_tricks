param(
    [string]$SourceFolder = (Get-Location)
)

$word = New-Object -ComObject Word.Application
$word.Visible = $true

# Ensure the folder exists
if (-Not (Test-Path $SourceFolder)) {
    Write-Error "The folder '$SourceFolder' does not exist."
    exit
}

# Get all .docx files (excluding temp/hidden)
$docxFiles = Get-ChildItem -Path $SourceFolder -Filter *.docx -File

foreach ($file in $docxFiles) {
    $fullPath = $file.FullName
    Write-Output "Processing: $fullPath"
    $pdfPath = [System.IO.Path]::ChangeExtension($fullPath, ".pdf")

    try {
        $doc = $word.Documents.Open($fullPath, [ref]$false, [ref]$true)
        $doc.SaveAs([ref]$pdfPath, [ref]17)  # 17 = wdFormatPDF
        $doc.Close()
        Write-Output "Converted: $file.Name to PDF"
    } catch {
        Write-Error "Failed to convert: $file.Name"
    }
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
