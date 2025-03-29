import subprocess
import os
import time
import ctypes
import sys
import tempfile
import shutil


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
    cmd = f'powershell -Command "Start-Process -FilePath \'{sys.executable}\' -ArgumentList \'{script}\' -Verb RunAs"'

    try:
        subprocess.run(cmd, shell=True)
    except Exception as e:
        print(f"Error while trying to elevate privileges: {e}")
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
        if "sfc" in command or "chkdsk" in command or "winget" in command or "DISM" in command:
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
    except Exception as e:
        print(f"\n❌ Error executing {description}: {str(e)}")
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
    with open(batch_path, "w") as batch_file:
        batch_file.write(batch_content)

    return batch_path


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

    # 2. DISM RestoreHealth (comprehensive repair)
    print("\n==== DISM SYSTEM IMAGE REPAIR ====")
    print("This operation may take 10-30 minutes to complete.")
    print("Please be patient while the system image is being repaired...")

    if run_command("DISM /Online /Cleanup-Image /RestoreHealth", "DISM Restore Health") == 0:
        success_count += 1

    # 3. Run System File Checker
    if run_command("sfc /scannow", "System File Checker") == 0:
        success_count += 1

    # 4. Reset Winsock
    if run_command("netsh winsock reset", "Reset Winsock Catalog") == 0:
        success_count += 1

    # 5. Reset TCP/IP Stack
    if run_command("netsh int ip reset", "Reset TCP/IP Stack") == 0:
        success_count += 1
        print("Note: A system restart is recommended after resetting TCP/IP stack.")

    # 6. Update all software with Winget
    if run_command(
            "winget update --all --include-unknown --accept-source-agreements --accept-package-agreements --silent",
            "Update all software with Winget"
    ) == 0:
        success_count += 1

    # 4. Check Disk
    print("\n==== DISK CHECK SCHEDULING ====")
    print("Check Disk requires a system restart to run on your system drive.")
    run_chkdsk = input("Do you want to schedule Check Disk to run on next restart? (y/n): ").lower()

    if run_chkdsk == 'y':
        if run_command("chkdsk C: /f /r", "Schedule Check Disk") == 0:
            success_count += 1
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
    print(f"MAINTENANCE SUMMARY: {success_count}/{total_commands} tasks completed successfully")
    print(f"{'=' * 60}")

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
</assembly>
'''
    manifest_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "elevation.manifest")
    try:
        with open(manifest_path, "w") as f:
            f.write(manifest)
        return manifest_path
    except Exception as e:
        print(f"Could not create manifest file: {e}")
        return None


if __name__ == "__main__":
    # You can uncomment the line below if you plan to compile this script to an exe
    # and want to use the manifest method instead
    # create_manifest_file()

    main()