# Add the following command to the windows task manager
#   "pwsh C:\Users\johan\Documents\Repositories\cmd_tricks\windows\time_track\track_time.ps1"

$storageFolder = $HOME + "/track_time"
If (!(test-path $storageFolder)) {
    New-Item -ItemType Directory -Force -Path $storageFolder
}

$category = "nothing"
If (test-path $storageFolder\category.txt) {
    $category = $(Get-Content $storageFolder/category.txt)
}

echo "<http://$(hostname)/time/timePoint/$([guid]::NewGuid().tostring())> <http://jvsoest.eu/ontology/timeTrack.owl#at_time> `"$((Get-Date).tostring("yyyy-MM-ddTHH:mm:ss"))`"^^<http://www.w3.org/2001/XMLSchema#dateTime>; <http://jvsoest.eu/ontology/timeTrack.owl#hostname> <http://$(hostname)>; <http://jvsoest.eu/ontology/timeTrack.owl#category> <http://$category/>." >> $HOME/track_time/$((Get-Date).tostring("yyyyMMdd")).ttl