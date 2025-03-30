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
    except Exception:
        return False


def run_as_admin():
    """Re-run the script with administrator privileges"""
    # If frozen to exe (using PyInstaller or similar)
    if getattr(sys, 'frozen', False):
        script = sys.executable
    else:
        script = os.path.abspath(sys.argv[0])

    # Command to elevate privileges
    cmd = (f'powershell -Command "Start-Process -FilePath \'{sys.executable}\' '
           f'-ArgumentList \'{script}\' -Verb RunAs"')

    try:
        subprocess.run(cmd, shell=True)
    except subprocess.SubprocessError as elevation_error:
        print(f"Error while trying to elevate privileges: {elevation_error}")
        input("Press Enter to exit...")

    # Exit the current non-elevated process
    sys.exit(0)


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
            process = subprocess.Popen(command, shell=True)

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


def clear_temp_files():
    """Clean temporary files and AppData folders"""
    cleaned_locations = []
    cleaned_files_count = 0
    cleaned_size_mb = 0

    print("\n==== CLEANING TEMPORARY FILES AND APPDATA ====")

    # Get system and user environment paths
    temp_dir = os.environ.get('TEMP')
    windows_dir = os.environ.get('SystemRoot', 'C:\\Windows')
    appdata_dir = os.environ.get('APPDATA')
    localappdata_dir = os.environ.get('LOCALAPPDATA')
    userprofile_dir = os.environ.get('USERPROFILE')

    # Locations to clean
    temp_locations = [
        # Windows Temp folders
        temp_dir,
        os.path.join(windows_dir, 'Temp'),
        # User's Temporary Internet Files
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\INetCache'),
        # Windows Prefetch folder
        os.path.join(windows_dir, 'Prefetch'),
        # Recent items
        os.path.join(appdata_dir, 'Microsoft\\Windows\\Recent'),
        # Windows Error Reporting
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\WER'),
        # Windows Update Cleanup
        os.path.join(windows_dir, 'SoftwareDistribution\\Download'),
    ]

    # AppData folders to clean
    appdata_locations = [
        # Browsers cache
        os.path.join(localappdata_dir, 'Google\\Chrome\\User Data\\Default\\Cache'),
        os.path.join(localappdata_dir, 'Mozilla\\Firefox\\Profiles'),
        os.path.join(localappdata_dir, 'Microsoft\\Edge\\User Data\\Default\\Cache'),
        # Application caches
        os.path.join(localappdata_dir, 'Microsoft\\Teams\\Cache'),
        os.path.join(localappdata_dir, 'Microsoft\\Teams\\blob_storage'),
        os.path.join(appdata_dir, 'Zoom\\logs'),
        os.path.join(appdata_dir, 'Slack\\Cache'),
        os.path.join(appdata_dir, 'Spotify\\Data'),
        os.path.join(localappdata_dir, 'Discord\\Cache'),
        os.path.join(localappdata_dir, 'Microsoft\\Office\\16.0\\OfficeFileCache'),
        # Windows app caches
        os.path.join(localappdata_dir, 'Packages'),
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\Explorer'),
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\Caches'),
        # Downloaded Program Files
        os.path.join(windows_dir, 'Downloaded Program Files'),
        # Temporary Downloads (optional - uncomment if desired)
        # os.path.join(userprofile_dir, 'Downloads'),
    ]

    # Combine all locations to clean
    all_locations = temp_locations + appdata_locations

    # File extensions to target for cleanup
    extensions_to_clean = ['.tmp', '.temp', '.log', '.old', '.bak', '.chk', '.~',
                           '.cache', '.etl', '.evt', '.dmp', '.dump']

    print(f"Scanning for files in {len(all_locations)} locations...")

    # Process each location
    for location in all_locations:
        if not location or not os.path.exists(location):
            continue

        try:
            print(f"\nCleaning: {location}")
            location_file_count = 0
            location_size_bytes = 0

            # If it's a directory, walk through and delete files
            if os.path.isdir(location):
                for root, dirs, files in os.walk(location, topdown=False):
                    # Delete files first
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            # Check if the file has an extension to clean
                            _, ext = os.path.splitext(file_path)
                            # Clean if it matches extensions or is in an appdata location
                            if ext.lower() in extensions_to_clean or location in appdata_locations:
                                try:
                                    file_size = os.path.getsize(file_path)
                                    os.remove(file_path)
                                    location_file_count += 1
                                    location_size_bytes += file_size
                                except (PermissionError, OSError):
                                    # Skip files that are in use or protected
                                    pass
                        except Exception:
                            # Skip files with problems
                            pass

                    # Try to remove empty folders in appdata locations
                    if location in appdata_locations:
                        for dir_name in dirs:
                            try:
                                dir_path = os.path.join(root, dir_name)
                                if os.path.exists(dir_path) and not os.listdir(dir_path):
                                    os.rmdir(dir_path)
                            except (PermissionError, OSError):
                                pass

            # Track statistics
            if location_file_count > 0:
                location_size_mb = location_size_bytes / (1024 * 1024)
                cleaned_locations.append(location)
                cleaned_files_count += location_file_count
                cleaned_size_mb += location_size_mb
                print(f"✓ Removed {location_file_count} files ({location_size_mb:.2f} MB)")
            else:
                print("✓ No applicable files to clean")

        except Exception as temp_error:
            print(f"× Error cleaning {location}: {str(temp_error)}")

    # Run Windows Disk Cleanup utility for system files
    print("\n==== SYSTEM DISK CLEANUP ====")
    run_disk_cleanup = input("Would you like to run the Windows Disk Cleanup utility for system files? (y/n): ").lower()
    if run_disk_cleanup == 'y':
        print("\nRunning Windows Disk Cleanup utility...")
        # Use /sageset:1 to run with pre-configured settings (if previously set up)
        # Or use /d to clean the system drive
        run_command("cleanmgr /d C:", "Windows Disk Cleanup")

    # Summary
    print(f"\n=== TEMPORARY FILES CLEANUP SUMMARY ===")
    print(f"Total files removed: {cleaned_files_count}")
    print(f"Total space freed: {cleaned_size_mb:.2f} MB")
    print(f"Locations cleaned: {len(cleaned_locations)}")

    return cleaned_files_count > 0


def clear_browser_data():
    """Clear browser history, cookies, and cache"""
    print("\n==== BROWSER DATA CLEANUP ====")
    print("This will clear browser history, cookies, and cached data.")
    print("WARNING: You will be signed out of websites and need to login again.")

    proceed = input("\nDo you want to clear browser data? (y/n): ").lower()
    if proceed != 'y':
        print("Browser data cleanup skipped.")
        return False

    browsers_cleared = []

    # Chrome
    try:
        chrome_profiles_path = os.path.join(os.environ.get('LOCALAPPDATA', ''),
                                            'Google\\Chrome\\User Data')
        if os.path.exists(chrome_profiles_path):
            print("\nClearing Google Chrome data...")
            # Find all profiles (Default and Profile*)
            profiles = ['Default']
            for item in os.listdir(chrome_profiles_path):
                if item.startswith('Profile '):
                    profiles.append(item)

            # Try to close Chrome first
            try:
                subprocess.run('taskkill /F /IM chrome.exe', shell=True,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(1)  # Give it time to close
            except subprocess.SubprocessError:
                pass

            # Clear data for each profile
            for profile in profiles:
                profile_path = os.path.join(chrome_profiles_path, profile)
                if os.path.isdir(profile_path):
                    # Locations to clear
                    locations = [
                        os.path.join(profile_path, 'History'),
                        os.path.join(profile_path, 'History-journal'),
                        os.path.join(profile_path, 'Cookies'),
                        os.path.join(profile_path, 'Cookies-journal'),
                        os.path.join(profile_path, 'Cache'),
                        os.path.join(profile_path, 'Media Cache'),
                        os.path.join(profile_path, 'Visited Links'),
                        os.path.join(profile_path, 'Network Action Predictor'),
                        os.path.join(profile_path, 'Login Data'),
                        os.path.join(profile_path, 'Web Data')
                    ]

                    for loc in locations:
                        try:
                            if os.path.isfile(loc):
                                os.remove(loc)
                            elif os.path.isdir(loc):
                                shutil.rmtree(loc, ignore_errors=True)
                        except (PermissionError, OSError):
                            pass

            browsers_cleared.append("Chrome")
            print("✓ Chrome data cleared")
    except Exception as chrome_error:
        print(f"× Error clearing Chrome data: {str(chrome_error)}")

    # Firefox
    try:
        firefox_path = os.path.join(os.environ.get('APPDATA', ''),
                                    'Mozilla\\Firefox\\Profiles')
        if os.path.exists(firefox_path):
            print("\nClearing Firefox data...")
            # Try to close Firefox first
            try:
                subprocess.run('taskkill /F /IM firefox.exe', shell=True,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(1)  # Give it time to close
            except subprocess.SubprocessError:
                pass

            # Process each profile
            for root, dirs, files in os.walk(firefox_path):
                for file in files:
                    if file in ['cookies.sqlite', 'cookies.sqlite-journal',
                                'places.sqlite', 'places.sqlite-journal',
                                'formhistory.sqlite', 'webappsstore.sqlite']:
                        try:
                            os.remove(os.path.join(root, file))
                        except (PermissionError, OSError):
                            pass

            browsers_cleared.append("Firefox")
            print("✓ Firefox data cleared")
    except Exception as firefox_error:
        print(f"× Error clearing Firefox data: {str(firefox_error)}")

    # Edge
    try:
        edge_path = os.path.join(os.environ.get('LOCALAPPDATA', ''),
                                 'Microsoft\\Edge\\User Data')
        if os.path.exists(edge_path):
            print("\nClearing Microsoft Edge data...")
            # Try to close Edge first
            try:
                subprocess.run('taskkill /F /IM msedge.exe', shell=True,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                time.sleep(1)  # Give it time to close
            except subprocess.SubprocessError:
                pass

            # Find all profiles (Default and Profile*)
            profiles = ['Default']
            for item in os.listdir(edge_path):
                if item.startswith('Profile '):
                    profiles.append(item)

            # Clear data for each profile
            for profile in profiles:
                profile_path = os.path.join(edge_path, profile)
                if os.path.isdir(profile_path):
                    # Locations to clear
                    locations = [
                        os.path.join(profile_path, 'History'),
                        os.path.join(profile_path, 'History-journal'),
                        os.path.join(profile_path, 'Cookies'),
                        os.path.join(profile_path, 'Cookies-journal'),
                        os.path.join(profile_path, 'Cache'),
                        os.path.join(profile_path, 'Media Cache'),
                        os.path.join(profile_path, 'Visited Links')
                    ]

                    for loc in locations:
                        try:
                            if os.path.isfile(loc):
                                os.remove(loc)
                            elif os.path.isdir(loc):
                                shutil.rmtree(loc, ignore_errors=True)
                        except (PermissionError, OSError):
                            pass

            browsers_cleared.append("Edge")
            print("✓ Edge data cleared")
    except Exception as edge_error:
        print(f"× Error clearing Edge data: {str(edge_error)}")

    # Summary
    if browsers_cleared:
        print(f"\n✅ Successfully cleared data for: {', '.join(browsers_cleared)}")
        print("Note: You will need to log in again to websites in these browsers.")
        return True
    else:
        print("\n⚠️ No browser data was cleared.")
        return False


def clean_registry():
    """Clean Windows Registry of common issues"""
    print("\n==== REGISTRY CLEANUP ====")
    print("This will scan and clean the Windows Registry for common issues.")
    print("WARNING: Registry cleaning can potentially cause system issues if critical entries are removed.")
    print("This tool only targets safe, common problem areas.")

    proceed = input("\nDo you want to clean the Windows Registry? (y/n): ").lower()
    if proceed != 'y':
        print("Registry cleanup skipped.")
        return False

    print("\nScanning registry for issues...")
    registry_fixes = 0

    # Create a temporary REG file with cleanup commands
    try:
        temp_reg_file = os.path.join(tempfile.gettempdir(), "registry_cleanup.reg")
        with open(temp_reg_file, "w") as reg_file:
            reg_file.write("Windows Registry Editor Version 5.00\n\n")

            # Remove broken/invalid file associations
            reg_file.write("; Remove broken file associations\n")
            reg_file.write(
                "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.???\\UserChoice]\n\n")

            # Clean MUICache (unused program entries)
            reg_file.write("; Clean MUICache\n")
            reg_file.write("[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\ShellNoRoam\\MUICache]\n")
            reg_file.write("[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\ShellNoRoam\\MUICache]\n\n")

            # Remove orphaned software entries
            reg_file.write("; Remove orphaned software entries\n")
            reg_file.write("[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit]\n")
            reg_file.write("[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit]\n\n")

            # Clean recent documents list
            reg_file.write("; Clean recent documents list\n")
            reg_file.write("[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs]\n")
            reg_file.write(
                "[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs]\n\n")

            # Clean Run MRU (Most Recently Used commands)
            reg_file.write("; Clean Run MRU list\n")
            reg_file.write("[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU]\n")
            reg_file.write("[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU]\n\n")

            # Clean typedURLs (URLs typed in Internet Explorer/Edge)
            reg_file.write("; Clean typed URLs\n")
            reg_file.write("[-HKEY_CURRENT_USER\\Software\\Microsoft\\Internet Explorer\\TypedURLs]\n")
            reg_file.write("[HKEY_CURRENT_USER\\Software\\Microsoft\\Internet Explorer\\TypedURLs]\n\n")

        # Execute the REG file
        print("Applying registry fixes...")
        result = subprocess.run(f'regedit /s "{temp_reg_file}"', shell=True)

        if result.returncode == 0:
            registry_fixes += 6  # Number of sections we cleaned
            print("✅ Registry fixes applied successfully.")
        else:
            print("⚠️ Some registry fixes may not have been applied.")

        # Clean up temporary file
        try:
            os.remove(temp_reg_file)
        except (PermissionError, OSError):
            pass

        # Run additional registry cleanup using PowerShell
        print("\nPerforming deep registry scan...")
        ps_cmd = 'powershell -Command "& {Get-ItemProperty -Path HKCU:\\Software\\*\\* -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -eq \'\' } | Remove-Item -ErrorAction SilentlyContinue -Force; exit 0}"'
        subprocess.run(ps_cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # Summary
        print(f"\n=== REGISTRY CLEANUP SUMMARY ===")
        print(f"Fixed registry entries: At least {registry_fixes}")
        print("Note: A system restart is recommended to fully apply registry changes.")

        return True

    except Exception as reg_error:
        print(f"⚠️ Error during registry cleanup: {str(reg_error)}")
        return False


def main():
    if not is_admin():
        print("This script requires administrator privileges.")
        print("Requesting elevation now...")

        # Method 1: Use PowerShell to relaunch
        run_as_admin()

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
    total_commands = 10

    # Track operations that require restart
    restart_required_ops = []

    # 1. Flush DNS Cache
    if run_command("ipconfig /flushdns", "Flush DNS Cache") == 0:
        success_count += 1

    # 2. Clean Temporary Files and AppData
    print("\nStarting temporary files and AppData cleanup...")
    if clear_temp_files():
        success_count += 1
        print("✅ Temporary files and AppData cleanup completed successfully.")
    else:
        print("⚠️ Temporary files cleanup completed with possible skipped items.")

    # 3. Clean Browser Data (Optional)
    if clear_browser_data():
        success_count += 1
        print("✅ Browser data cleanup completed successfully.")
    else:
        print("Browser data cleanup skipped.")
        total_commands -= 1  # Adjust total if skipped

    # 4. Clean Registry (Optional)
    if clean_registry():
        success_count += 1
        restart_required_ops.append("Registry Cleanup")
        print("✅ Registry cleanup completed successfully.")
    else:
        print("Registry cleanup skipped.")
        total_commands -= 1  # Adjust total if skipped

    # 5. DISM System Image Check and Repair
    print("\n==== DISM SYSTEM IMAGE CHECK ====")
    print("First performing a quick scan to check if repairs are needed...")

    # Run the quicker ScanHealth first
    scan_result = run_command("DISM /Online /Cleanup-Image /ScanHealth", "DISM Scan Health")

    # Check if corruption was found and run RestoreHealth only if needed
    if scan_result == 0:
        # Need to check output to see if corruption was detected
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

    # 6. Run System File Checker
    if run_command("sfc /scannow", "System File Checker") == 0:
        success_count += 1

    # 7. Reset Winsock
    if run_command("netsh winsock reset", "Reset Winsock Catalog") == 0:
        success_count += 1
        restart_required_ops.append("Winsock Reset")

    # 8. Reset TCP/IP Stack
    if run_command("netsh int ip reset", "Reset TCP/IP Stack") == 0:
        success_count += 1
        restart_required_ops.append("TCP/IP Stack Reset")
        print("Note: A system restart is required for network stack changes to take effect.")

    # 9. Update all software with Winget
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
                                except ValueError:
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

    # 10. Check Disk
    print("\n==== DISK CHECK SCHEDULING ====")
    print("Check Disk requires a system restart to run on your system drive.")
    run_chkdsk = input("Do you want to schedule Check Disk to run on next restart? (y/n): ").lower()

    if run_chkdsk == 'y':
        if run_command("chkdsk C: /f /r", "Schedule Check Disk") == 0:
            success_count += 1
            restart_required_ops.append("Check Disk")
            print("\nCheck Disk has been scheduled to run on the next system restart.")
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


if __name__ == "__main__":
    # Create manifest file for UAC elevation when compiled to exe
    create_manifest_file()

    main()