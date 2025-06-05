$exclude = @("venv", "RelatorioBT.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "RelatorioBT.zip" -Force