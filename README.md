The script will automatically check if it is run as admin, otherwise it prompts admin rights.

The following powershell prompts will be run.

ipconfig /flushdns  # Clear DNS cache

winget update --all --include-unknown --accept-source-agreements --accept-package-agreements --silent  # Update all apps

sfc /scannow # Scan and repair windows system files

chkdsk C: /f /r  # Check and repair disk errors
