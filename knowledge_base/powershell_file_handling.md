# Powershell file handling

This KB contains several examples for file handling using PowerShell. Every sub-heading contains a specific example.

## Loop over folder, and create zip file

In this example, we are looping (Foreach) over a folder (using Get-ChildItem). The collected file/folder names are replaced by using a regular expression, and this variable is afterwards used when zipping (compressing) the folder into a zip file.

```
Foreach ($folder in Get-ChildItem) {
    $objectName = $folder.name -replace '(.*)Session', ''
    Rename-Item $folder.name $objectName
    Compress-Archive -Path $objectName/* -DestinationPath $objectName'.zip'
    Remove-Item -Recurse $objectName
}
```
