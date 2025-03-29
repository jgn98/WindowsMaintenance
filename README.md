A Python script that automatically runs common Windows maintenance commands with admin privileges.
Commands Executed

ipconfig /flushdns - Clears DNS cache
DISM /Online /Cleanup-Image /RestoreHealth - Repairs Windows system image
sfc /scannow - Scans and repairs system files
netsh winsock reset - Resets Windows Sockets
netsh int ip reset - Resets TCP/IP stack
winget update --all --include-unknown --accept-source-agreements --accept-package-agreements --silent - Updates all software
chkdsk C: /f /r - Schedules disk check (optional)

Usage

Save as pc_maintenance.py
Run with python pc_maintenance.py
Admin rights will be requested automatically
Follow on-screen prompts

Requirements

Windows 10/11
Python 3.6+
Winget package manager

Notes

Some tasks (DISM) may take 15+ minutes
System restart recommended after running
TCP/IP reset may temporarily disconnect network

