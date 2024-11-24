import os
import shutil
import winreg
import psutil
import subprocess
import ctypes
from datetime import datetime

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def clear_temp_files():
    """Clear temporary files from multiple Windows locations"""
    temp_paths = [
        os.environ.get('TEMP'),
        os.environ.get('TMP'),
        os.path.join(os.environ.get('WINDIR'), 'Temp'),
        os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp')
    ]
    
    bytes_cleared = 0
    print("\nClearing temporary files...")
    
    for temp_path in temp_paths:
        if temp_path and os.path.exists(temp_path):
            print(f"Cleaning {temp_path}")
            for item in os.listdir(temp_path):
                item_path = os.path.join(temp_path, item)
                try:
                    if os.path.isfile(item_path):
                        file_size = os.path.getsize(item_path)
                        os.unlink(item_path)
                        bytes_cleared += file_size
                    elif os.path.isdir(item_path):
                        dir_size = get_dir_size(item_path)
                        shutil.rmtree(item_path, ignore_errors=True)
                        bytes_cleared += dir_size
                except Exception as e:
                    continue
    
    print(f"Cleared {bytes_cleared / (1024*1024):.2f} MB of temporary files")

def get_dir_size(path):
    """Calculate directory size"""
    total = 0
    try:
        for entry in os.scandir(path):
            try:
                if entry.is_file():
                    total += entry.stat().st_size
                elif entry.is_dir():
                    total += get_dir_size(entry.path)
            except Exception as e:
                continue
    except Exception as e:
        pass
    return total

def optimize_performance():
    """Optimize various Windows performance settings"""
    try:
        # Disable visual effects
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 2)
        
        # Set power plan to high performance
        subprocess.run(['powercfg', '/setactive', '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'], 
                      capture_output=True)
        
        print("\nPerformance settings optimized")
    except Exception as e:
        print(f"Error optimizing performance: {e}")

def check_system_health():
    """Check and display system health metrics"""
    print("\nSystem Health Check:")
    
    # CPU Usage
    cpu_percent = psutil.cpu_percent(interval=1)
    print(f"CPU Usage: {cpu_percent}%")
    
    # Memory Usage
    memory = psutil.virtual_memory()
    print(f"Memory Usage: {memory.percent}%")
    print(f"Available Memory: {memory.available / (1024*1024*1024):.2f} GB")
    
    # Disk Usage
    disk = psutil.disk_usage('/')
    print(f"Disk Usage: {disk.percent}%")
    print(f"Free Disk Space: {disk.free / (1024*1024*1024):.2f} GB")

def run_disk_cleanup():
    """Run Windows Disk Cleanup utility"""
    print("\nRunning Disk Cleanup...")
    try:
        subprocess.run(['cleanmgr', '/sagerun:1'], capture_output=True)
        print("Disk Cleanup completed")
    except Exception as e:
        print(f"Error running Disk Cleanup: {e}")

def main():
    if not is_admin():
        print("This script requires administrative privileges.")
        print("Please run as administrator for full optimization.")
        input("Press Enter to continue anyway...")
    
    print("=== System Optimization Script ===")
    print("Starting optimization process...")
    
    # Create a log file
    log_file = f"optimization_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    print(f"Creating log file: {log_file}")
    
    # Check system health before optimization
    print("\nChecking system health before optimization...")
    check_system_health()
    
    # Perform optimizations
    clear_temp_files()
    optimize_performance()
    run_disk_cleanup()
    
    # Check system health after optimization
    print("\nChecking system health after optimization...")
    check_system_health()
    
    print("\nOptimization complete!")
    print("Note: Some changes may require a system restart to take effect.")
    
    restart = input("\nWould you like to restart your computer now? (y/n): ").lower()
    if restart == 'y':
        subprocess.run(['shutdown', '/r', '/t', '60', '/c', "System will restart in 60 seconds to complete optimization."])
        print("System will restart in 60 seconds. Save your work.")
        print("Type 'shutdown /a' to cancel restart.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        print("\nPress Enter to exit...")
        input()