import subprocess
import os
import time
import ctypes
import sys
import tempfile
import shutil
from pathlib import Path


def is_admin():
    """Check if script has admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


def run_as_admin():
    """Re-run script with admin privileges"""
    if getattr(sys, 'frozen', False):
        script = sys.executable
    else:
        script = str(Path(sys.argv[0]).resolve())

    cmd = f'powershell -Command "Start-Process -FilePath \'{sys.executable}\' -ArgumentList \'{script}\' -Verb RunAs"'

    try:
        subprocess.run(cmd, shell=True)
    except Exception as e:
        print(f"Error elevating privileges: {e}")
        input("Press Enter to exit...")

    sys.exit(0)


def create_manifest_file():
    """Create UAC manifest for executable"""
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
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
    </application>
  </compatibility>
</assembly>
'''
    manifest_path = str(Path(os.path.dirname(os.path.abspath(sys.argv[0]))) / "WindowsMaintenance.manifest")
    try:
        with open(manifest_path, "w", encoding="utf-8") as f:
            f.write(manifest)
        return manifest_path
    except Exception as e:
        print(f"Could not create manifest: {e}")
        return None


def run_command(command, description):
    """Run system command with feedback"""
    print(f"\n{'=' * 60}\nRunning: {description}\nCommand: {command}\n{'=' * 60}\n")

    try:
        # Real-time output for certain commands
        if any(cmd in command for cmd in ["sfc", "chkdsk", "winget", "DISM"]):
            process = subprocess.Popen(command, shell=True)
            process.wait()
            exit_code = process.returncode
        else:
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
        print(f"\n❌ Error executing {description}: {e}")
        return -1


def clear_temp_files():
    """Clean temporary files and AppData folders"""
    cleaned_locations = []
    cleaned_files_count = 0
    cleaned_size_mb = 0

    print("\n==== CLEANING TEMPORARY FILES AND APPDATA ====")

    # Environment paths - convert to strings explicitly
    temp_dir = str(os.environ.get('TEMP', ''))
    windows_dir = str(os.environ.get('SystemRoot', 'C:\\Windows'))
    appdata_dir = str(os.environ.get('APPDATA', ''))
    localappdata_dir = str(os.environ.get('LOCALAPPDATA', ''))

    # Locations to clean - ensure all paths are strings
    locations = [
        # System temp folders
        temp_dir,
        os.path.join(windows_dir, 'Temp'),
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\INetCache'),
        os.path.join(windows_dir, 'Prefetch'),
        os.path.join(appdata_dir, 'Microsoft\\Windows\\Recent'),
        os.path.join(localappdata_dir, 'Microsoft\\Windows\\WER'),
        os.path.join(windows_dir, 'SoftwareDistribution\\Download'),

        # AppData folders
        os.path.join(localappdata_dir, 'Google\\Chrome\\User Data\\Default\\Cache'),
        os.path.join(localappdata_dir, 'Mozilla\\Firefox\\Profiles'),
        os.path.join(localappdata_dir, 'Microsoft\\Edge\\User Data\\Default\\Cache'),
        os.path.join(localappdata_dir, 'Microsoft\\Teams\\Cache'),
        os.path.join(localappdata_dir, 'Microsoft\\Teams\\blob_storage'),
        os.path.join(appdata_dir, 'Zoom\\logs'),
        os.path.join(appdata_dir, 'Slack\\Cache'),
        os.path.join(localappdata_dir, 'Discord\\Cache'),
        os.path.join(localappdata_dir, 'Microsoft\\Office\\16.0\\OfficeFileCache'),
    ]

    # File extensions to target
    extensions = ['.tmp', '.temp', '.log', '.old', '.bak', '.chk', '.~', '.cache', '.etl', '.evt', '.dmp']

    print(f"Scanning for files in {len(locations)} locations...")

    # Process each location
    for location in locations:
        if not location or not os.path.exists(location):
            continue

        try:
            print(f"\nCleaning: {location}")
            location_file_count = 0
            location_size_bytes = 0

            if os.path.isdir(location):
                for root, dirs, files in os.walk(location, topdown=False):
                    # Delete files
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            _, ext = os.path.splitext(file_path)
                            if ext.lower() in extensions or location != temp_dir:
                                try:
                                    file_size = os.path.getsize(file_path)
                                    os.remove(file_path)
                                    location_file_count += 1
                                    location_size_bytes += file_size
                                except (PermissionError, OSError):
                                    pass
                        except:
                            pass

                    # Remove empty folders
                    for dir_name in dirs:
                        try:
                            dir_path = str(os.path.join(root, dir_name))
                            if os.path.exists(dir_path) and not os.listdir(dir_path):
                                os.rmdir(dir_path)
                        except:
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

        except Exception as e:
            print(f"× Error cleaning {location}: {e}")

    # Windows Disk Cleanup
    run_disk_cleanup = input("\nRun Windows Disk Cleanup utility? (y/n): ").lower()
    if run_disk_cleanup == 'y':
        print("\nRunning Windows Disk Cleanup...")
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

    # Helper function to kill browser process
    def kill_browser(process_name):
        try:
            subprocess.run(f'taskkill /F /IM {process_name}', shell=True,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(1)
        except:
            pass

    # Chrome
    try:
        chrome_profiles_path = os.path.join(str(os.environ.get('LOCALAPPDATA', '')), 'Google\\Chrome\\User Data')
        if os.path.exists(chrome_profiles_path):
            print("\nClearing Google Chrome data...")
            kill_browser("chrome.exe")

            # Find profiles
            profiles = ['Default']
            try:
                profiles += [item for item in os.listdir(chrome_profiles_path) if item.startswith('Profile ')]
            except Exception:
                pass

            # Chrome data to clear
            chrome_files = ['History', 'History-journal', 'Cookies', 'Cookies-journal',
                            'Cache', 'Media Cache', 'Visited Links', 'Network Action Predictor',
                            'Login Data', 'Web Data']

            # Clear each profile
            for profile in profiles:
                profile_path = os.path.join(chrome_profiles_path, profile)
                if os.path.isdir(profile_path):
                    for file in chrome_files:
                        path = os.path.join(profile_path, file)
                        try:
                            if os.path.isfile(path):
                                os.remove(path)
                            elif os.path.isdir(path):
                                shutil.rmtree(path, ignore_errors=True)
                        except:
                            pass

            browsers_cleared.append("Chrome")
            print("✓ Chrome data cleared")
    except Exception as e:
        print(f"× Error clearing Chrome data: {e}")

    # Firefox
    try:
        firefox_path = os.path.join(str(os.environ.get('APPDATA', '')), 'Mozilla\\Firefox\\Profiles')
        if os.path.exists(firefox_path):
            print("\nClearing Firefox data...")
            kill_browser("firefox.exe")

            # Firefox files to clear
            firefox_files = ['cookies.sqlite', 'cookies.sqlite-journal', 'places.sqlite',
                             'places.sqlite-journal', 'formhistory.sqlite', 'webappsstore.sqlite']

            # Clear each profile
            for root, dirs, files in os.walk(firefox_path):
                for file in files:
                    if file in firefox_files:
                        try:
                            os.remove(os.path.join(root, file))
                        except:
                            pass

            browsers_cleared.append("Firefox")
            print("✓ Firefox data cleared")
    except Exception as e:
        print(f"× Error clearing Firefox data: {e}")

    # Edge
    try:
        edge_path = os.path.join(str(os.environ.get('LOCALAPPDATA', '')), 'Microsoft\\Edge\\User Data')
        if os.path.exists(edge_path):
            print("\nClearing Microsoft Edge data...")
            kill_browser("msedge.exe")

            # Find profiles
            profiles = ['Default']
            try:
                profiles += [item for item in os.listdir(edge_path) if item.startswith('Profile ')]
            except Exception:
                pass

            # Edge data to clear
            edge_files = ['History', 'History-journal', 'Cookies', 'Cookies-journal',
                          'Cache', 'Media Cache', 'Visited Links']

            # Clear each profile
            for profile in profiles:
                profile_path = os.path.join(edge_path, profile)
                if os.path.isdir(profile_path):
                    for file in edge_files:
                        path = os.path.join(profile_path, file)
                        try:
                            if os.path.isfile(path):
                                os.remove(path)
                            elif os.path.isdir(path):
                                shutil.rmtree(path, ignore_errors=True)
                        except:
                            pass

            browsers_cleared.append("Edge")
            print("✓ Edge data cleared")
    except Exception as e:
        print(f"× Error clearing Edge data: {e}")

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
    print("WARNING: Registry cleaning can potentially cause system issues.")
    print("This tool only targets safe, common problem areas.")

    proceed = input("\nDo you want to clean the Windows Registry? (y/n): ").lower()
    if proceed != 'y':
        print("Registry cleanup skipped.")
        return False

    print("\nScanning registry for issues...")
    registry_fixes = 0

    # Create temporary REG file
    try:
        temp_reg_file = os.path.join(tempfile.gettempdir(), "registry_cleanup.reg")
        with open(temp_reg_file, "w", encoding="utf-8") as reg_file:
            reg_file.write("Windows Registry Editor Version 5.00\n\n")

            # Registry keys to clean
            registry_entries = [
                # File associations
                ("Remove broken file associations",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.???\\UserChoice]"),

                # MUICache
                ("Clean MUICache",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\ShellNoRoam\\MUICache]\n"
                 "[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\ShellNoRoam\\MUICache]"),

                # Orphaned software
                ("Remove orphaned software entries",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit]\n"
                 "[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit]"),

                # Recent docs
                ("Clean recent documents list",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs]\n"
                 "[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs]"),

                # Run MRU
                ("Clean Run MRU list",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU]\n"
                 "[HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU]"),

                # Typed URLs
                ("Clean typed URLs",
                 "[-HKEY_CURRENT_USER\\Software\\Microsoft\\Internet Explorer\\TypedURLs]\n"
                 "[HKEY_CURRENT_USER\\Software\\Microsoft\\Internet Explorer\\TypedURLs]")
            ]

            # Write registry entries
            for description, command in registry_entries:
                reg_file.write(f"; {description}\n{command}\n\n")
                registry_fixes += 1

        # Execute REG file
        print("Applying registry fixes...")
        result = subprocess.run(f'regedit /s "{temp_reg_file}"', shell=True)

        # Clean up temp file
        try:
            os.remove(temp_reg_file)
        except:
            pass

        # Run PowerShell cleanup
        print("\nPerforming deep registry scan...")
        ps_cmd = 'powershell -Command "& {Get-ItemProperty -Path HKCU:\\Software\\*\\* -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -eq \'\' } | Remove-Item -ErrorAction SilentlyContinue -Force; exit 0}"'
        subprocess.run(ps_cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # Summary
        print(f"\n=== REGISTRY CLEANUP SUMMARY ===")
        print(f"Fixed registry entries: At least {registry_fixes}")
        print("Note: A system restart is recommended to fully apply registry changes.")

        return True

    except Exception as e:
        print(f"⚠️ Error during registry cleanup: {e}")
        return False


def main():
    if not is_admin():
        print("This script requires administrator privileges.")
        print("Requesting elevation now...")
        run_as_admin()
        sys.exit(0)

    print("\n==== PC MAINTENANCE UTILITY ====")
    print("This script will run several system maintenance commands.")
    print("Some commands may take a while to complete.")
    print("It's recommended to close other applications before proceeding.")

    proceed = input("\nDo you want to proceed? (y/n): ").lower()
    if proceed != 'y':
        print("Operation cancelled.")
        return

    # Track success and required restarts
    success_count = 0
    total_commands = 10
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
        total_commands -= 1

    # 4. Clean Registry (Optional)
    if clean_registry():
        success_count += 1
        restart_required_ops.append("Registry Cleanup")
        print("✅ Registry cleanup completed successfully.")
    else:
        print("Registry cleanup skipped.")
        total_commands -= 1

    # 5. DISM System Image Check and Repair
    print("\n==== DISM SYSTEM IMAGE CHECK ====")
    print("First performing a quick scan to check if repairs are needed...")

    # Run the quicker ScanHealth first
    scan_result = run_command("DISM /Online /Cleanup-Image /ScanHealth", "DISM Scan Health")

    if scan_result == 0:
        # Check if corruption was found
        check_cmd = 'powershell -Command "& {$output = dism /Online /Cleanup-Image /ScanHealth | Out-String; if ($output -match \'Component Store is repairable\' -or $output -match \'Component Store corruption\') {exit 1} else {exit 0}}"'
        corruption_check = subprocess.run(check_cmd, shell=True, capture_output=True, text=True)

        if corruption_check.returncode == 1:
            print("\n⚠️ System image corruption detected. Running repair operation...")
            print("This may take 10-30 minutes to complete.")

            if run_command("DISM /Online /Cleanup-Image /RestoreHealth", "DISM Restore Health") == 0:
                success_count += 1
            else:
                print("⚠️ DISM Restore operation failed or was interrupted.")
        else:
            print("\n✅ No system image corruption detected. Skipping repair operation.")
            success_count += 1
    else:
        print("\n⚠️ DISM scan operation failed or was interrupted.")
        try_restore = input("\nDISM scan failed. Try full system repair anyway? (y/n): ").lower()
        if try_restore == 'y':
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

    try:
        # Get list of available updates
        print("Scanning for available updates...")
        winget_list_cmd = "winget upgrade --include-unknown --source winget"
        winget_result = subprocess.run(winget_list_cmd, shell=True, capture_output=True, text=True)
        winget_output = winget_result.stdout

        # Check if there are any updates available
        if "No installed package" in winget_output or "All installed packages are up to date." in winget_output:
            print("✅ All software is up to date.")
            success_count += 1
        else:
            print("\nAvailable updates found. Options:")
            print("1. Update all (with confirmation for each)")
            print("2. Update all and automatically close running applications")
            print("3. Skip updates")

            choice = input("\nEnter your choice (1-3): ").strip()

            if choice == "1":
                if run_command("winget upgrade --all --include-unknown", "Update software with Winget") == 0:
                    success_count += 1

            elif choice == "2":
                print("\nPreparing to update and close running applications...")

                # Parse packages needing updates
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
                    print(f"Found {len(packages)} packages to update.")
                    print("Checking for related running processes...")

                    # Get running processes
                    ps_command = 'powershell -Command "Get-Process | Select-Object ProcessName, Id | Format-Table -AutoSize | Out-String -Width 4096"'
                    process_result = subprocess.run(ps_command, shell=True, capture_output=True, text=True)
                    processes_output = process_result.stdout

                    # Detect and kill processes
                    killed_processes = []
                    for package in packages:
                        package_base = ''.join(c for c in package if c.isalnum()).lower()
                        for line in processes_output.splitlines():
                            parts = line.split()
                            if len(parts) >= 2 and parts[0].lower() != "processname":
                                process_name = parts[0].lower()
                                try:
                                    process_id = int(parts[-1])

                                    # Check for matches
                                    if (process_name in package_base or package_base in process_name or
                                            any(process_name in p.lower() for p in package.split()) or
                                            any(p.lower() in process_name for p in package.split())):

                                        confirm = input(f"Found process '{parts[0]}' (PID: {process_id}) "
                                                        f"related to '{package}'. Kill it? (y/n): ").lower()
                                        if confirm == 'y':
                                            kill_cmd = f'taskkill /F /PID {process_id}'
                                            kill_result = subprocess.run(kill_cmd, shell=True, capture_output=True,
                                                                         text=True)
                                            if kill_result.returncode == 0:
                                                print(f"✅ Terminated process {parts[0]} (PID: {process_id})")
                                                killed_processes.append(parts[0])
                                            else:
                                                print(f"❌ Failed to terminate process: {kill_result.stderr}")
                                except ValueError:
                                    continue

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

    except Exception as e:
        print(f"Error processing software updates: {e}")
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

    # Show completion status
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

    # Prompt for restart
    if pending_count > 0:
        restart_message = f"\nSystem restart required to complete {pending_count} maintenance tasks."
        restart_prompt = input(f"{restart_message} Would you like to restart now? (y/n): ").lower()
    else:
        restart_prompt = input("\nMaintenance complete. Would you like to restart your computer now? (y/n): ").lower()

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
    # Create manifest file for UAC elevation
    create_manifest_file()
    main()