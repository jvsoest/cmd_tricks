Foreach ($item in Get-ChildItem $HOME\Documents\Repositories\cmd_tricks\windows) {
    If (!($item -is [System.IO.DirectoryInfo])) {
        $item.name
    }
}