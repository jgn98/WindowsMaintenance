import subprocess
import os
import time
import ctypes
import sys
import tempfile


def is_admin():
    """Check if the script is running with administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


def run_as_admin():
    """Re-run the script with administrator privileges"""
    # If frozen to exe (using PyInstaller or similar)
    if getattr(sys, 'frozen', False):
        script = sys.executable
    else:
        script = os.path.abspath(sys.argv[0])

    # Command to elevate privileges
    params = ' '.join([f'"{item}"' for item in sys.argv[1:]])

    # The command that will be run with elevated privileges
    cmd = (f'powershell -Command "Start-Process -FilePath \'{sys.executable}\' '
           f'-ArgumentList \'{script}\' -Verb RunAs"')

    try:
        subprocess.run(cmd, shell=True)
    except subprocess.SubprocessError as elevation_error:
        print(f"Error while trying to elevate privileges: {elevation_error}")
        input("Press Enter to exit...")

    # Exit the current non-elevated process
    sys.exit(0)


def run_command(command, description):
    """Run a system command and provide feedback"""
    print(f"\n{'=' * 60}")
    print(f"Running: {description}")
    print(f"Command: {command}")
    print(f"{'=' * 60}\n")

    try:
        # For commands that need to show their output in real-time
        if any(cmd in command for cmd in ["sfc", "chkdsk", "winget", "DISM"]):
            # Run with direct console output instead of capturing
            process = subprocess.Popen(
                command,
                shell=True,
                # Don't redirect stdout/stderr, let them go directly to console
                # This prevents the character-by-character output issue
            )

            # Wait for the process to complete
            process.wait()
            exit_code = process.returncode
        else:
            # For simpler commands
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            print(result.stdout)
            if result.stderr:
                print("Error:", result.stderr)
            exit_code = result.returncode

        if exit_code == 0:
            print(f"\n✅ {description} completed successfully.")
        else:
            print(f"\n❌ {description} failed with exit code {exit_code}.")

        return exit_code
    except subprocess.SubprocessError as subprocess_error:
        print(f"\n❌ Error executing {description}: {str(subprocess_error)}")
        return -1


def create_admin_batch_launcher():
    """Create a batch file that will run the script with admin privileges"""
    # Get the full path of the current script
    if getattr(sys, 'frozen', False):
        script_path = sys.executable
    else:
        script_path = os.path.abspath(sys.argv[0])

    # Create a temporary directory for the batch file
    temp_dir = tempfile.mkdtemp()
    batch_path = os.path.join(temp_dir, "run_with_admin.bat")

    # Create batch file content
    batch_content = f"""@echo off
echo Requesting administrative privileges...
powershell -Command "Start-Process -FilePath '{sys.executable}' -ArgumentList '{script_path}' -Verb RunAs"
"""

    # Write the batch file
    try:
        with open(batch_path, "w") as batch_file:
            batch_file.write(batch_content)
        return batch_path
    except OSError as file_error:
        print(f"Error creating batch file: {file_error}")
        return None


def main():
    if not is_admin():
        print("This script requires administrator privileges.")
        print("Requesting elevation now...")

        # Method 1: Use PowerShell to relaunch
        run_as_admin()

        # Method 2 (alternative): Create and run a batch file
        # Uncomment below if Method 1 doesn't work
        # batch_path = create_admin_batch_launcher()
        # subprocess.Popen(batch_path, shell=True)

        # Exit current non-elevated process
        sys.exit(0)

    print("\n==== PC MAINTENANCE UTILITY ====")
    print("This script will run several system maintenance commands.")
    print("Some commands may take a while to complete.")
    print("It's recommended to close other applications before proceeding.")

    proceed = input("\nDo you want to proceed? (y/n): ").lower()
    if proceed != 'y':
        print("Operation cancelled.")
        return

    # Track overall success
    success_count = 0
    total_commands = 7

    # 1. Flush DNS Cache
    if run_command("ipconfig /flushdns", "Flush DNS Cache") == 0:
        success_count += 1

    # 2. DISM System Image Check and Repair
    print("\n==== DISM SYSTEM IMAGE CHECK ====")
    print("First performing a quick scan to check if repairs are needed...")

    # Run the quicker ScanHealth first
    scan_result = run_command("DISM /Online /Cleanup-Image /ScanHealth", "DISM Scan Health")

    # Check if corruption was found and run RestoreHealth only if needed
    if scan_result == 0:
        # Need to check output to see if corruption was detected
        # This requires capturing the output from ScanHealth
        check_cmd = 'powershell -Command "& {$output = dism /Online /Cleanup-Image /ScanHealth | Out-String; if ($output -match \'Component Store is repairable\' -or $output -match \'Component Store corruption\') {exit 1} else {exit 0}}"'
        corruption_check = subprocess.run(check_cmd, shell=True, capture_output=True, text=True)

        if corruption_check.returncode == 1:
            print("\n⚠️ System image corruption detected. Running repair operation...")
            print("This may take 10-30 minutes to complete.")
            print("Please be patient while the system image is being repaired...")

            if run_command("DISM /Online /Cleanup-Image /RestoreHealth", "DISM Restore Health") == 0:
                success_count += 1
            else:
                print("⚠️ DISM Restore operation failed or was interrupted.")
        else:
            print("\n✅ No system image corruption detected. Skipping repair operation.")
            success_count += 1
    else:
        print("\n⚠️ DISM scan operation failed or was interrupted.")

        # Ask user if they want to try RestoreHealth anyway
        try_restore = input("\nDISM scan failed. Would you like to try full system repair anyway? (y/n): ").lower()
        if try_restore == 'y':
            print("\nRunning system image repair. This may take 10-30 minutes...")
            if run_command("DISM /Online /Cleanup-Image /RestoreHealth", "DISM Restore Health") == 0:
                success_count += 1

    # 3. Run System File Checker
    if run_command("sfc /scannow", "System File Checker") == 0:
        success_count += 1

    # Track operations that require restart
    restart_required_ops = []

    # 4. Reset Winsock
    if run_command("netsh winsock reset", "Reset Winsock Catalog") == 0:
        success_count += 1
        restart_required_ops.append("Winsock Reset")

    # 5. Reset TCP/IP Stack
    if run_command("netsh int ip reset", "Reset TCP/IP Stack") == 0:
        success_count += 1
        restart_required_ops.append("TCP/IP Stack Reset")
        print("Note: A system restart is required for network stack changes to take effect.")

    # 6. Update all software with Winget (interactive mode with process termination)
    print("\n==== SOFTWARE UPDATES ====")
    print("Checking for software updates with Winget...")

    # First get list of available updates
    print("Scanning for available updates...")
    winget_list_cmd = "winget upgrade --include-unknown --source winget"

    try:
        # Get the update list
        winget_result = subprocess.run(winget_list_cmd, shell=True, capture_output=True, text=True)
        winget_output = winget_result.stdout

        # Check if there are any updates available
        if ("No installed package" in winget_output
                or "All installed packages are up to date." in winget_output):
            print("✅ All software is up to date.")
            success_count += 1
        else:
            print("\nAvailable updates found. Would you like to:")
            print("1. Update all (with confirmation for each)")
            print("2. Update all and automatically close running applications")
            print("3. Skip updates")

            choice = input("\nEnter your choice (1-3): ").strip()

            if choice == "1":
                # Regular update with confirmation
                if run_command("winget upgrade --all --include-unknown", "Update software with Winget") == 0:
                    success_count += 1

            elif choice == "2":
                print("\nPreparing to update and close running applications...")

                # Parse the winget output to get list of packages that need updating
                packages = []
                for line in winget_output.splitlines():
                    if "winget" in line and "upgrade" in line:
                        continue
                    if "|" in line and "..." not in line and "Name" not in line and "----" not in line:
                        parts = line.split("|")
                        if len(parts) >= 2:
                            package_name = parts[0].strip()
                            if package_name and package_name != "Name":
                                packages.append(package_name)

                if packages:
                    # Function to find and kill processes that might be related to these packages
                    print(f"Found {len(packages)} packages to update.")
                    print("Checking for related running processes...")

                    # Get list of running processes
                    ps_command = 'powershell -Command "Get-Process | Select-Object ProcessName, Id | Format-Table -AutoSize | Out-String -Width 4096"'
                    process_result = subprocess.run(ps_command, shell=True, capture_output=True, text=True)
                    processes_output = process_result.stdout

                    # Try to detect and kill processes
                    killed_processes = []
                    for package in packages:
                        # Create variations of the package name to match against processes
                        # Strip non-alphanumeric to get base name
                        package_base = ''.join(c for c in package if c.isalnum()).lower()
                        for line in processes_output.splitlines():
                            parts = line.split()
                            if len(parts) >= 2 and parts[0].lower() != "processname":
                                process_name = parts[0].lower()
                                try:
                                    process_id = int(parts[-1])

                                    # Check for potential matches with package name
                                    if (process_name in package_base or package_base in process_name or
                                            any(process_name in p.lower() for p in package.split()) or
                                            any(p.lower() in process_name for p in package.split())):

                                        confirm = input(
                                            f"Found running process '{parts[0]}' (PID: {process_id}) "
                                            f"that may be related to '{package}'. Kill it? (y/n): "
                                        ).lower()
                                        if confirm == 'y':
                                            # Kill the process
                                            kill_cmd = f'taskkill /F /PID {process_id}'
                                            kill_result = subprocess.run(kill_cmd, shell=True, capture_output=True,
                                                                         text=True)
                                            if kill_result.returncode == 0:
                                                print(
                                                    f"✅ Successfully terminated process {parts[0]} (PID: {process_id})")
                                                killed_processes.append(parts[0])
                                            else:
                                                print(f"❌ Failed to terminate process: {kill_result.stderr}")
                                except:
                                    # Skip if process ID is not a number
                                    continue

                    # Now run the upgrades
                    if killed_processes:
                        print(f"\nTerminated {len(killed_processes)} processes: {', '.join(killed_processes)}")

                    print("\nProceeding with software updates...")
                    if run_command("winget upgrade --all --include-unknown", "Update software with Winget") == 0:
                        success_count += 1
                else:
                    print("No packages identified for update.")

            elif choice == "3":
                print("Skipping software updates.")

            else:
                print("Invalid choice. Skipping software updates.")

    except Exception as specific_error:
        print(f"Error processing software updates: {str(specific_error)}")
        print("Proceeding with standard update method.")
        if run_command("winget upgrade --all --include-unknown", "Update software with Winget") == 0:
            success_count += 1

    # 4. Check Disk
    print("\n==== DISK CHECK SCHEDULING ====")
    print("Check Disk requires a system restart to run on your system drive.")
    run_chkdsk = input("Do you want to schedule Check Disk to run on next restart? (y/n): ").lower()

    if run_chkdsk == 'y':
        if run_command("chkdsk C: /f /r", "Schedule Check Disk") == 0:
            success_count += 1
            restart_required_ops.append("Check Disk")
            print("\nCheck Disk has been scheduled to run on the next system restart.")
            restart_now = input("Would you like to restart now? (y/n): ").lower()
            if restart_now == 'y':
                print("System will restart in 10 seconds...")
                time.sleep(10)
                os.system("shutdown /r /t 0")
    else:
        print("Check Disk operation skipped.")
        total_commands -= 1

    # Final summary
    print(f"\n\n{'=' * 60}")

    # Show more accurate completion status
    pending_count = len(restart_required_ops)
    immediate_success = success_count - pending_count

    print(f"MAINTENANCE SUMMARY:")
    print(f"• {immediate_success}/{total_commands} tasks completed successfully")

    if pending_count > 0:
        print(f"• {pending_count} tasks scheduled and will complete after restart:")
        for op in restart_required_ops:
            print(f"  - {op}")

    failed_count = total_commands - success_count
    if failed_count > 0:
        print(f"• {failed_count} tasks did not complete successfully")

    print(f"{'=' * 60}")

    # Prompt for system restart
    if pending_count > 0:
        restart_message = f"\nSystem restart required to complete {pending_count} maintenance tasks."
        restart_prompt = input(f"{restart_message} Would you like to restart now? (y/n): ").lower()
    else:
        restart_prompt = input("\nMaintenance complete. Would you like to restart your computer "
                               "to apply all changes? (y/n): ").lower()

    if restart_prompt == 'y':
        print("\nSaving any remaining work and restarting in 10 seconds...")
        print("Close this window to cancel restart.")
        time.sleep(10)
        os.system("shutdown /r /t 0")
    else:
        if pending_count > 0:
            print("\n⚠️ Please restart your computer soon to complete the pending maintenance tasks.")
        else:
            print("\nPlease consider restarting your computer at your earliest convenience.")

    input("\nPress Enter to exit...")


def create_manifest_file():
    """Create a manifest file for UAC elevation"""
    # This is an alternative method that works when compiled to exe
    manifest = '''
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
  <application xmlns="urn:schemas-microsoft-com:asm.v3">
    <windowsSettings>
      <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>
      <longPathAware xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">true</longPathAware>
    </windowsSettings>
  </application>
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <!-- Windows 10 and 11 -->
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
</assembly>
'''
    manifest_path = os.path.join(
        os.path.dirname(os.path.abspath(sys.argv[0])),
        "WindowsMaintenance.manifest"
    )
    try:
        with open(manifest_path, "w") as f:
            f.write(manifest)
        return manifest_path
    except OSError as manifest_error:
        print(f"Could not create manifest file: {manifest_error}")
        return None


if __name__ == "__main__":
    # Create manifest file for UAC elevation when compiled to exe
    create_manifest_file()

    main()