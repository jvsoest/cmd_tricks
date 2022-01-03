Foreach ($item in Get-ChildItem $HOME\Repositories\cmd_tricks\windows) {
    If (!($item -is [System.IO.DirectoryInfo])) {
        $item.name
    }
}
