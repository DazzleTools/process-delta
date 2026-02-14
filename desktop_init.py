# Attempted to adapt pyWinVirtualDesktop https://github.com/kdschlosser/pyWinVirtualDesktop
# Local functioning build, but still issues with portability and ease of use
# So instead using built-in Virtual Desktop detection via DWM cloaking with pywin32

from __future__ import print_function
import os
import subprocess
import time
import ctypes
from ctypes import wintypes
import psutil
import sys
import argparse

# Try to import win32 modules, but make them optional
try:
    import win32api, win32gui, win32process, win32com.client, win32con
    WIN32_AVAILABLE = True
except ImportError:
    print("Warning: pywin32 not available. Using limited functionality.")
    WIN32_AVAILABLE = False

# DWM constants for virtual desktop detection
DWMWA_CLOAKED = 14

class WindowInfo:
    """Simple container for window information"""
    def __init__(self, hwnd, title, process_name):
        self.id = hwnd
        self.text = title
        self.process_name = process_name
        self.is_on_active_desktop = True  # Will be set by detection

class VirtualDesktopDetector:
    """Handles virtual desktop detection using DWM cloaking"""
    
    def __init__(self):
        self.dwmapi = ctypes.WinDLL("dwmapi")
        self.user32 = ctypes.WinDLL("user32")
        self.kernel32 = ctypes.WinDLL("kernel32")
        self.psapi = ctypes.WinDLL("psapi")
        
    def is_window_on_current_desktop(self, hwnd):
        """Check if window is on current virtual desktop using DWM cloaking"""
        try:
            # Check if window is visible
            if not self.user32.IsWindowVisible(hwnd):
                return False
            
            # Check cloaked state
            cloaked = ctypes.c_int(0)
            result = self.dwmapi.DwmGetWindowAttribute(
                wintypes.HWND(hwnd),
                wintypes.DWORD(DWMWA_CLOAKED),
                ctypes.byref(cloaked),
                ctypes.sizeof(cloaked)
            )
            
            # Window is on current desktop if not cloaked
            return result == 0 and cloaked.value == 0
        except:
            # Fallback - assume visible windows are on current desktop
            return self.user32.IsWindowVisible(hwnd)
    
    def get_window_text(self, hwnd):
        """Get window title"""
        length = self.user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return ""
        
        buffer = ctypes.create_unicode_buffer(length + 1)
        self.user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value
    
    def get_process_name_from_hwnd(self, hwnd):
        """Get process name from window handle"""
        try:
            # Get process ID
            pid = wintypes.DWORD()
            self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            
            # Open process
            PROCESS_QUERY_INFORMATION = 0x0400
            PROCESS_VM_READ = 0x0010
            handle = self.kernel32.OpenProcess(
                PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                False,
                pid.value
            )
            
            if handle:
                # Get process name
                filename = ctypes.create_unicode_buffer(260)  # MAX_PATH
                if self.psapi.GetModuleBaseNameW(handle, None, filename, 260):
                    self.kernel32.CloseHandle(handle)
                    return filename.value
                self.kernel32.CloseHandle(handle)
        except:
            pass
        
        return "Unknown"
    
    def enumerate_desktop_windows(self):
        """Enumerate all windows on current desktop"""
        windows = []
        
        def enum_handler(hwnd, param):
            # Skip windows without titles
            title = self.get_window_text(hwnd)
            if not title:
                return True
            
            # Check if on current desktop
            if self.is_window_on_current_desktop(hwnd):
                process_name = self.get_process_name_from_hwnd(hwnd)
                window = WindowInfo(hwnd, title, process_name)
                window.is_on_active_desktop = True
                windows.append(window)
            
            return True
        
        # Define the callback type
        WNDENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool,
            ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(ctypes.c_int)
        )
        
        # Enumerate windows
        self.user32.EnumWindows(WNDENUMPROC(enum_handler), 0)
        
        return windows

class SimpleDesktop:
    """Mimics the pyWinVirtualDesktop desktop interface"""
    def __init__(self):
        self.id = "current"
        self.is_active = True
        self._detector = VirtualDesktopDetector()
    
    def __iter__(self):
        """Iterate through windows on this desktop"""
        return iter(self._detector.enumerate_desktop_windows())

class FallbackDesktop:
    """Fallback when DWM detection isn't available"""
    def __init__(self):
        self.id = "current"
        self.is_active = True
    
    def __iter__(self):
        """Use psutil to enumerate processes with windows"""
        windows = []
        
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # Skip processes without names
                if not proc.info['name']:
                    continue
                
                # Create a fake window entry for each process
                # This is less accurate but works as fallback
                window = WindowInfo(
                    proc.info['pid'],
                    proc.info['name'],
                    proc.info['name']
                )
                windows.append(window)
            except:
                continue
        
        return iter(windows)

# Create shell object globally if available
shell = None
if WIN32_AVAILABLE:
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
    except:
        pass

# Track launched shortcuts to allow proper duplicate handling
launched_shortcuts = set()

# Global configuration for multiple instances
allow_multiple_default = True
restricted_programs = set()

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description='Desktop Startup Script - Launch shortcuts with duplicate control',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s
    # Default: Allow multiple instances of all programs
    
  %(prog)s --restrict-multiple firefox chrome
    # Restrict Firefox and Chrome to single instance
    
  %(prog)s -rm firefox -rm "Google Chrome" -rm edge
    # Restrict multiple programs (use quotes for names with spaces)
    
  %(prog)s --restrict-all
    # Old behavior: Restrict all programs by default
    
  %(prog)s --restrict-all --allow-multiple notepad cmd
    # Restrict all except notepad and cmd
    
  %(prog)s "C:\\Custom\\Startup\\Folder"
    # Use a custom startup folder
        """
    )
    
    parser.add_argument('startup_dir', nargs='?', default=None,
                        help='Startup directory path (default: ./Desktop-Startup)')
    
    # Multiple instance control
    parser.add_argument('--restrict-all', action='store_true',
                        help='Restrict all programs to single instance by default')
    
    parser.add_argument('-rm', '--restrict-multiple', action='append',
                        dest='restrict_list', metavar='PROGRAM',
                        help='Restrict specific program to single instance (can be used multiple times)')
    
    parser.add_argument('-am', '--allow-multiple', action='append',
                        dest='allow_list', metavar='PROGRAM',
                        help='Allow multiple instances of specific program (only useful with --restrict-all)')
    
    # Other options
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Show detailed output')
    
    parser.add_argument('--delay', type=float, default=1.0,
                        help='Delay between launching programs (default: 1.0 seconds)')
    
    parser.add_argument('--wait-time', type=int, default=5,
                        help='Maximum time to wait for program to start (default: 5 seconds)')
    
    args = parser.parse_args()
    
    # Process the arguments into global configuration
    global allow_multiple_default, restricted_programs
    
    if args.restrict_all:
        allow_multiple_default = False
        # If restrict_all is set, allow_list specifies exceptions
        if args.allow_list:
            # These programs will be allowed multiple instances
            # We'll handle this by NOT adding them to restricted_programs
            pass
    else:
        allow_multiple_default = True
        # If not restrict_all, restrict_list specifies what to restrict
        if args.restrict_list:
            restricted_programs = set(prog.lower() for prog in args.restrict_list)
    
    # Handle the combination of --restrict-all and --allow-multiple
    if args.restrict_all and args.allow_list:
        # In this case, everything is restricted EXCEPT what's in allow_list
        # We'll store the allow_list separately
        args.allowed_programs = set(prog.lower() for prog in args.allow_list)
    else:
        args.allowed_programs = None
    
    return args

def should_allow_multiple(filename, shortcut_targetname, args):
    """Determine if multiple instances should be allowed for this program"""
    
    # Extract the base name without extension
    target_base = shortcut_targetname.lower()
    if target_base.endswith('.exe'):
        target_base = target_base[:-4]
    
    shortcut_base = os.path.splitext(filename)[0].lower()
    
    # Check special markers in filename
    if '--' in filename or ' - ' in filename:
        return True
    
    # If restrict_all mode with allow_list
    if args.restrict_all and args.allowed_programs:
        # Check if this program is in the allowed list
        for allowed in args.allowed_programs:
            if allowed in target_base or allowed in shortcut_base:
                return True
        return False
    
    # If restrict_all mode without allow_list
    if args.restrict_all:
        return False
    
    # Default mode: check if program is in restricted list
    for restricted in restricted_programs:
        if restricted in target_base or restricted in shortcut_base:
            return False
    
    # Default behavior
    return allow_multiple_default

def IsShortcutAlreadyRunning(filename, args):
    """Check if a shortcut's target is already running"""
    shortcut_basename = os.path.splitext(filename)[0].lower()
    
    # Special handling for already launched shortcuts in this session
    # This allows multiple instances of the same program with different shortcuts
    if filename.lower() in launched_shortcuts:
        return True
    
    # Try to use virtual desktop detection
    try:
        desktop = SimpleDesktop()
    except:
        if args.verbose:
            print("Using fallback process detection...")
        desktop = FallbackDesktop()
    
    # Check if we have win32com for proper shortcut parsing
    if shell:
        try:
            shortcut_obj = shell.CreateShortCut(filename)
            shortcut_targetname = os.path.basename(shortcut_obj.Targetpath.lower())
            shortcut_targetname_noext = os.path.splitext(shortcut_targetname)[0]
            shortcut_arguments = shortcut_obj.Arguments
        except:
            shortcut_targetname = shortcut_basename + ".exe"
            shortcut_targetname_noext = shortcut_basename
            shortcut_arguments = ""
    else:
        shortcut_targetname = shortcut_basename + ".exe"
        shortcut_targetname_noext = shortcut_basename
        shortcut_arguments = ""
    
    # Check if multiple instances are allowed for this program
    allow_multiple = should_allow_multiple(filename, shortcut_targetname, args)
    
    if args.verbose:
        print(f"  Multiple instances allowed: {allow_multiple}")
    
    # If multiple processes are allowed, check for exact shortcut name match
    if allow_multiple:
        # Only check if this specific shortcut variant is running
        for window in desktop:
            if window.is_on_active_desktop:
                window_title = window.text.lower()
                # Check if the window title contains the shortcut's unique identifier
                if shortcut_basename in window_title:
                    if args.verbose:
                        print(f"  Found matching window for {shortcut_basename}: {window.text}")
                    return True
        return False
    
    # Standard check for single-instance programs
    for window in desktop:
        if window.is_on_active_desktop:
            process_name = str(window.process_name).lower()
            
            # Remove .exe extension for comparison
            if process_name.endswith('.exe'):
                process_name = process_name[:-4]
            
            if (shortcut_basename in process_name or 
               shortcut_basename in shortcut_arguments or
               shortcut_targetname_noext in process_name):
                if args.verbose:
                    print(f"  Found matching process: {window.process_name}")
                return True
    
    return False

def main():
    """Main execution"""
    # Parse command line arguments
    args = parse_arguments()
    
    print("Desktop Startup Script - Windows 11 Compatible Version")
    print("=" * 50)
    
    # IMPORTANT: Use the current working directory (where the shortcut was launched)
    # NOT the script's directory
    initial_cwd = os.getcwd()
    print(f"Launched from: {initial_cwd}")
    
    # Determine startup directory
    if args.startup_dir:
        # If argument provided, use it
        if args.startup_dir == ".":
            startup_dir = os.path.join(initial_cwd, "Desktop-Startup")
        elif os.path.isabs(args.startup_dir):
            startup_dir = args.startup_dir
        else:
            startup_dir = os.path.join(initial_cwd, args.startup_dir)
    else:
        # Default: Look for Desktop-Startup in current working directory
        startup_dir = os.path.join(initial_cwd, "Desktop-Startup")
    
    # If startup_dir doesn't exist, try just using the current directory
    if not os.path.exists(startup_dir):
        print(f"Desktop-Startup not found at: {startup_dir}")
        print(f"Checking current directory for .lnk files...")
        
        # Check if current directory has .lnk files
        lnk_files = [f for f in os.listdir(initial_cwd) if f.endswith('.lnk')]
        if lnk_files:
            print(f"Found {len(lnk_files)} .lnk files in current directory")
            startup_dir = initial_cwd
        else:
            # Try creating Desktop-Startup folder
            print(f"Creating Desktop-Startup directory: {startup_dir}")
            os.makedirs(startup_dir, exist_ok=True)
    
    # Change to the startup directory
    os.chdir(startup_dir)
    print(f"Working directory: {os.getcwd()}")
    
    # Display configuration
    print(f"\nConfiguration:")
    print(f"  Multiple instances by default: {allow_multiple_default}")
    if restricted_programs:
        print(f"  Restricted programs: {', '.join(sorted(restricted_programs))}")
    if args.restrict_all and args.allowed_programs:
        print(f"  Allowed programs: {', '.join(sorted(args.allowed_programs))}")
    print(f"  Launch delay: {args.delay} seconds")
    print(f"  Max wait time: {args.wait_time} seconds")
    
    # Process shortcuts
    shortcuts = [f for f in os.listdir('.') if f.endswith('.lnk')]
    
    if not shortcuts:
        print("\nNo .lnk files found")
        print("\nTo use this script:")
        print("1. Create a 'Desktop-Startup' folder in your desired location")
        print("2. Place .lnk shortcuts in that folder")
        print("3. Set the shortcut's 'Start in' to the parent folder")
        print("   OR pass the path as an argument")
        return
    
    print(f"\nFound {len(shortcuts)} shortcuts to process")
    
    # Group shortcuts by their target to detect intentional duplicates
    shortcut_targets = {}
    if shell:
        for shortcut in shortcuts:
            try:
                obj = shell.CreateShortCut(shortcut)
                target = os.path.basename(obj.Targetpath.lower())
                if target not in shortcut_targets:
                    shortcut_targets[target] = []
                shortcut_targets[target].append(shortcut)
            except:
                pass
    
    # Report duplicate targets
    for target, files in shortcut_targets.items():
        if len(files) > 1:
            print(f"Note: {len(files)} shortcuts for {target}: {', '.join(files)}")
    
    print("-" * 50)
    
    # Clear launched shortcuts at start of each run
    launched_shortcuts.clear()
    
    for filename in shortcuts:
        basename = os.path.splitext(filename)[0]
        print(f"\nProcessing: {basename}")
        
        if not IsShortcutAlreadyRunning(filename, args):
            print(f"  Launching: {basename}")
            arguments = ['explorer.exe', filename]
            
            try:
                proc = subprocess.Popen(arguments, shell=False, stdin=None, 
                                      stdout=None, stderr=None, close_fds=True)
                
                # Mark as launched
                launched_shortcuts.add(filename.lower())
                
                # Wait for process to start
                counter = 0
                
                while counter < args.wait_time:
                    counter += 1
                    time.sleep(1)
                    
                    # For multiple instance programs, just wait a bit
                    if should_allow_multiple(filename, basename, args):
                        print(f"  Waiting for {basename}... ({counter}/{args.wait_time})")
                        if counter >= 2:  # Shorter wait for known multi-instance programs
                            print(f"  ✓ {basename} launched (multi-instance program)")
                            break
                    else:
                        # Check if process started for single-instance programs
                        if IsShortcutAlreadyRunning(filename, args):
                            print(f"  ✓ {basename} started successfully")
                            break
                        print(f"  Waiting for {basename} to start... ({counter}/{args.wait_time})")
                
                if counter >= args.wait_time:
                    print(f"  ! {basename} launched but window not detected (may be starting slowly)")
                
                time.sleep(args.delay)  # Configurable pause between launches
                
            except Exception as e:
                print(f"  ✗ Error launching {basename}: {e}")
        else:
            print(f"  → Skipping {basename} (already running)")
    
    print("\nDesktop initialization complete!")
    print(f"Launched {len(launched_shortcuts)} new shortcuts")

if __name__ == "__main__":
    main()