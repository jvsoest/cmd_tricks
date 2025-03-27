rem run vim in WSL from a cmd or PowerShell path
set filePath="%1"

set linuxSlash="/"
call set filePath=%%filePath:\=%linuxSlash%%%

bash -c "vim %filePath%"