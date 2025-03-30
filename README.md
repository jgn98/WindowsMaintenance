Windows Maintenance Utility
A comprehensive system maintenance and optimization tool for Windows. This utility automates common maintenance tasks to help keep your system running smoothly.
Features

Administrative Tasks

Auto-elevates privileges with UAC prompt
Creates proper manifest for executable


System Cleanup

Temporary files cleanup
AppData folder optimization
Browser data cleaning (history, cookies, cache)
Registry cleanup for common issues


System Repair

DISM system image verification and repair
System File Checker (SFC)
Disk error checking (CHKDSK)


Network Optimization

DNS cache flush
Winsock catalog reset
TCP/IP stack reset


Software Management

Automatic software updates with winget
Process termination for in-use applications



Usage

Run the executable. The tool will automatically request administrator privileges.
Select which maintenance tasks you want to perform.
Review the maintenance summary when complete.
Restart your computer when prompted to complete tasks that require a restart.

Commands Executed

ipconfig /flushdns - Clears DNS cache
Temporary Files Cleanup - Removes temp files from system and application folders
Browser Data Cleanup (optional) - Clears browser history, cookies, and cache
Registry Cleanup (optional) - Fixes common registry issues
DISM /Online /Cleanup-Image /RestoreHealth - Repairs Windows system image
sfc /scannow - Scans and repairs system files
netsh winsock reset - Resets Windows Sockets
netsh int ip reset - Resets TCP/IP stack
winget upgrade --all --include-unknown - Updates all software
chkdsk C: /f /r - Schedules disk check (optional)
