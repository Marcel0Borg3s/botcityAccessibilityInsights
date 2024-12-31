$exclude = @("venv", "bank.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "bank.zip" -Force