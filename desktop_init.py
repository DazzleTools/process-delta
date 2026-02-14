# Trying to get pyWinVirtualDesktop working: https://claude.ai/chat/7aa5f9ea-3bf2-4fff-8ccd-df85afad0ffd
# Final fixes: https://claude.ai/chat/b7bdf45d-c14d-4dcd-8add-2afd69c14fe2
# Updated to support native files (.ahk, .exe, .bat, etc.) in addition to .lnk shortcuts

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

# Define supported executable extensions
EXECUTABLE_EXTENSIONS = {'.exe', '.bat', '.cmd', '.ahk', '.ps1', '.vbs', '.com'}

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
        description='Desktop Startup Script - Launch shortcuts and executables with duplicate control',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s
    # Default: Allow multiple instances, include both .lnk and native files
    
  %(prog)s --no-native
    # Only process .lnk shortcuts, skip native executables
    
  %(prog)s --native-only
    # Only run native files (.exe, .ahk, .bat, etc.), skip .lnk shortcuts

  %(prog)s --native-types .ahk
    # Process .lnk shortcuts + only .ahk native files (skip .ps1, .bat, etc.)

  %(prog)s --native-types .ahk .exe
    # Process .lnk shortcuts + only .ahk and .exe native files

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
    
    # Native file handling
    parser.add_argument('--include-native', action='store_true',
                        help='Include native executable files in addition to .lnk shortcuts (this is the DEFAULT behavior)')
    
    parser.add_argument('--no-native', action='store_true',
                        help='Exclude native executable files, only process .lnk shortcuts')
    
    parser.add_argument('--native-only', action='store_true',
                        help='Only process native executable files, skip .lnk shortcuts')

    parser.add_argument('--native-types', nargs='+', metavar='EXT',
                        help='Only include these native file types (e.g., --native-types .ahk .exe). '
                             'When specified, only native files matching these extensions are processed. '
                             'Dot prefix is optional (both "ahk" and ".ahk" work).')
    
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
    
    # Handle native file inclusion
    # Simple: Default to True, only False if --no-native is used
    args.include_native = not args.no_native

    # Normalize --native-types extensions (ensure dot prefix, lowercase)
    if args.native_types:
        args.native_types = {
            ext.lower() if ext.startswith('.') else f'.{ext.lower()}'
            for ext in args.native_types
        }

    return args

def is_native_executable(filename):
    """Check if a file is a native executable"""
    # Make sure we're dealing with a file, not a directory
    if os.path.isdir(filename):
        return False
    
    # Get extension and normalize to lowercase
    ext = os.path.splitext(filename)[1].lower()
    
    # Strip any whitespace that might be present
    ext = ext.strip()
    
    return ext in EXECUTABLE_EXTENSIONS

def get_target_info(filename, args):
    """Get target information for either .lnk or native files"""
    basename = os.path.splitext(filename)[0].lower()
    
    if filename.endswith('.lnk'):
        if not shell:
            # Can't parse shortcuts without win32com
            if args.verbose:
                print(f"  Warning: Cannot parse shortcut '{filename}' - pywin32 not installed")
            return None, None, None
        
        try:
            shortcut_obj = shell.CreateShortCut(filename)
            targetname = os.path.basename(shortcut_obj.Targetpath.lower())
            targetname_noext = os.path.splitext(targetname)[0]
            arguments = shortcut_obj.Arguments
            return targetname, targetname_noext, arguments
        except Exception as e:
            # Shortcut parsing failed
            print(f"  Error: Failed to parse shortcut '{filename}': {e}")
            return None, None, None
    
    # For native files
    if is_native_executable(filename):
        targetname = filename.lower()
        targetname_noext = basename
        arguments = ""
        return targetname, targetname_noext, arguments
    
    # Unknown file type (shouldn't happen with current filtering)
    return None, None, None

def should_allow_multiple(filename, shortcut_targetname, args):
    """Determine if multiple instances should be allowed for this program"""
    
    # Handle case where we couldn't determine the target
    if shortcut_targetname is None:
        # Default to the global setting when we can't parse the file
        return allow_multiple_default
    
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

def IsFileAlreadyRunning(filename, args):
    """Check if a file's target is already running (works for both .lnk and native files)"""
    basename = os.path.splitext(filename)[0].lower()
    
    # Special handling for already launched files in this session
    if filename.lower() in launched_shortcuts:
        return True
    
    # Try to use virtual desktop detection
    try:
        desktop = SimpleDesktop()
    except:
        if args.verbose:
            print("Using fallback process detection...")
        desktop = FallbackDesktop()
    
    # Get target information
    targetname, targetname_noext, arguments = get_target_info(filename, args)
    
    # If we couldn't parse the file, we can't check if it's running
    if targetname is None:
        if args.verbose:
            print(f"  Cannot check if '{filename}' is running - unable to parse file")
        return False
    
    # Check if multiple instances are allowed for this program
    allow_multiple = should_allow_multiple(filename, targetname, args)
    
    if args.verbose:
        print(f"  Multiple instances allowed: {allow_multiple}")
    
    # If multiple processes are allowed, check for exact name match
    if allow_multiple:
        # Only check if this specific variant is running
        for window in desktop:
            if window.is_on_active_desktop:
                window_title = window.text.lower()
                # Check if the window title contains the file's unique identifier
                if basename in window_title:
                    if args.verbose:
                        print(f"  Found matching window for {basename}: {window.text}")
                    return True
        return False
    
    # Standard check for single-instance programs
    for window in desktop:
        if window.is_on_active_desktop:
            process_name = str(window.process_name).lower()
            
            # Remove .exe extension for comparison
            if process_name.endswith('.exe'):
                process_name = process_name[:-4]
            
            # For AHK files, check for AutoHotkey.exe
            if filename.endswith('.ahk'):
                if 'autohotkey' in process_name:
                    # Check window title for script name
                    # AutoHotkey typically includes the script name in the window title
                    window_title = window.text.lower()
                    
                    # Check if this specific script is running
                    # Look for the script filename in the window title
                    if basename in window_title or filename.lower() in window_title:
                        if args.verbose:
                            print(f"  Found matching AHK script: {window.text}")
                        return True
                    
                    # Also check if the process command line contains our script
                    # (This would require more advanced process inspection)
                    # For now, we'll be conservative and not match generic AutoHotkey processes
                    
                # Also check if there's a window with the script name
                # (Some AHK scripts create their own windows)
                elif basename in window.text.lower():
                    if args.verbose:
                        print(f"  Found window potentially from AHK script: {window.text}")
                    return True
            else:
                if (basename in process_name or 
                   basename in arguments or
                   targetname_noext in process_name):
                    if args.verbose:
                        print(f"  Found matching process: {window.process_name}")
                    return True
    
    return False

def launch_file(filename, args):
    """Launch either a .lnk shortcut or a native executable"""
    if filename.endswith('.lnk'):
        # Launch shortcut via explorer
        arguments = ['explorer.exe', filename]
    elif is_native_executable(filename):
        # Launch native file directly
        ext = os.path.splitext(filename)[1].lower()
        
        if ext == '.ahk':
            # AutoHotkey files - try multiple approaches
            # First, try to find AutoHotkey in common locations
            ahk_paths = [
                'AutoHotkey.exe',  # In PATH
                'AutoHotkeyU64.exe',  # v1 64-bit
                'AutoHotkeyU32.exe',  # v1 32-bit
                'AutoHotkey32.exe',  # v2 32-bit
                'AutoHotkey64.exe',  # v2 64-bit
                r'C:\Program Files\AutoHotkey\AutoHotkey.exe',
                r'C:\Program Files\AutoHotkey\v2\AutoHotkey.exe',
                r'C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe',
                r'C:\Program Files\AutoHotkey\v1.1\AutoHotkeyU64.exe',
                r'C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe',
            ]
            
            # Try each potential AutoHotkey executable
            ahk_found = False
            for ahk_exe in ahk_paths:
                try:
                    # Test if this AutoHotkey executable exists/works
                    arguments = [ahk_exe, filename]
                    proc = subprocess.Popen(arguments, shell=False, stdin=None,
                                          stdout=None, stderr=None, close_fds=True)
                    ahk_found = True
                    if args.verbose:
                        print(f"  Using AutoHotkey: {ahk_exe}")
                    return True
                except (FileNotFoundError, OSError):
                    continue
            
            if not ahk_found:
                # Last resort: try to launch .ahk directly (relies on Windows file association)
                try:
                    arguments = ['cmd.exe', '/c', 'start', '""', filename]
                    proc = subprocess.Popen(arguments, shell=False, stdin=None,
                                          stdout=None, stderr=None, close_fds=True)
                    if args.verbose:
                        print(f"  Launching via Windows file association")
                    return True
                except Exception as e:
                    print(f"  ✗ Error: AutoHotkey not found. Please install AutoHotkey or check PATH")
                    print(f"     Tried locations: {', '.join(ahk_paths[:5])}...")
                    return False
            
        elif ext in {'.bat', '.cmd'}:
            # Batch files - run via cmd
            arguments = ['cmd.exe', '/c', 'start', '""', filename]
        elif ext == '.ps1':
            # PowerShell scripts
            arguments = ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', filename]
        elif ext == '.vbs':
            # VBScript files
            arguments = ['wscript.exe', filename]
        else:
            # .exe, .com, and others - run directly
            arguments = [filename]
    else:
        # Unknown file type
        return False
    
    try:
        proc = subprocess.Popen(arguments, shell=False, stdin=None, 
                              stdout=None, stderr=None, close_fds=True)
        return True
    except Exception as e:
        print(f"  ✗ Error launching {filename}: {e}")
        return False

def main():
    """Main execution"""
    # Parse command line arguments
    args = parse_arguments()
    
    print("Desktop Startup Script - Windows 11 Compatible Version")
    print("(Now with native file support!)")
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
        print(f"Checking current directory for startup files...")
        
        # Check if current directory has relevant files
        all_files = os.listdir(initial_cwd)
        lnk_files = [f for f in all_files if f.endswith('.lnk')]
        native_files = [f for f in all_files if is_native_executable(f)]
        
        if lnk_files or native_files:
            print(f"Found {len(lnk_files)} .lnk files and {len(native_files)} native executables")
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
    print(f"  Include native files: {args.include_native}")
    if args.native_types:
        print(f"  Native types filter: {', '.join(sorted(args.native_types))}")
    print(f"  Native only: {args.native_only if hasattr(args, 'native_only') else False}")
    print(f"  Launch delay: {args.delay} seconds")
    print(f"  Max wait time: {args.wait_time} seconds")
    if args.verbose:
        print(f"  Recognized native extensions: {', '.join(sorted(EXECUTABLE_EXTENSIONS))}")
    
    # Collect files to process
    all_files = []
    
    # Debug: Show all files in directory first
    if args.verbose:
        print(f"\n=== Directory Contents ===")
        try:
            all_items = os.listdir('.')
            print(f"Total items in directory: {len(all_items)}")
            for item in sorted(all_items):
                full_path = os.path.join('.', item)
                if os.path.isdir(full_path):
                    print(f"  [DIR]  {item}")
                else:
                    size = os.path.getsize(full_path)
                    print(f"  [FILE] {item} ({size} bytes)")
        except Exception as e:
            print(f"Error listing directory: {e}")
        print("=" * 30)
    
    if not args.native_only:
        # Add .lnk files
        shortcuts = [f for f in os.listdir('.') if f.endswith('.lnk')]
        all_files.extend(shortcuts)
        print(f"\nFound {len(shortcuts)} .lnk shortcuts")
        if args.verbose and shortcuts:
            print(f"  Shortcuts: {', '.join(sorted(shortcuts))}")
    
    if args.include_native:
        # Add native executable files
        native_files = [f for f in os.listdir('.') if is_native_executable(f)]
        # Apply --native-types filter if specified
        if args.native_types:
            skipped = [f for f in native_files if os.path.splitext(f)[1].lower() not in args.native_types]
            native_files = [f for f in native_files if os.path.splitext(f)[1].lower() in args.native_types]
            if skipped:
                print(f"  Filtered by --native-types {' '.join(sorted(args.native_types))}:")
                print(f"    Skipped {len(skipped)}: {', '.join(sorted(skipped))}")
        all_files.extend(native_files)
        print(f"Found {len(native_files)} native executables")
        if native_files:
            if args.verbose:
                print(f"  Native files: {', '.join(sorted(native_files))}")
                # Extra debug: check each native file extension
                for nf in sorted(native_files):
                    ext = os.path.splitext(nf)[1].lower()
                    print(f"    - {nf} (ext: '{ext}')")
        else:
            # Debug: Let's see what files were skipped
            if args.verbose:
                print("  No native executables found. Checking what was skipped...")
                for f in os.listdir('.'):
                    if not f.endswith('.lnk') and os.path.isfile(f):
                        ext = os.path.splitext(f)[1].lower()
                        is_exec = ext in EXECUTABLE_EXTENSIONS
                        print(f"    - {f} (ext: '{ext}', is_executable: {is_exec})")
    else:
        if args.verbose:
            print("\nNative files DISABLED (--no-native flag used)")
            # Show what native files would have been processed
            native_files = [f for f in os.listdir('.') if is_native_executable(f)]
            if native_files:
                print(f"  Skipping {len(native_files)} native files: {', '.join(sorted(native_files))}")
    
    if not all_files:
        print("\nNo startup files found")
        print("\nTo use this script:")
        print("1. Create a 'Desktop-Startup' folder in your desired location")
        print("2. Place .lnk shortcuts and/or executable files in that folder")
        print("3. Set the shortcut's 'Start in' to the parent folder")
        print("   OR pass the path as an argument")
        return
    
    print(f"\nTotal files to process: {len(all_files)}")
    
    # Group files by their target to detect intentional duplicates
    file_targets = {}
    unparseable_shortcuts = []
    for file in all_files:
        targetname, _, _ = get_target_info(file, args)
        if targetname is None and file.endswith('.lnk'):
            unparseable_shortcuts.append(file)
            continue
        elif targetname:
            target = targetname.lower()
            if target not in file_targets:
                file_targets[target] = []
            file_targets[target].append(file)
    
    # Report duplicate targets
    for target, files in file_targets.items():
        if len(files) > 1:
            print(f"Note: {len(files)} files for {target}: {', '.join(files)}")
    
    # Report unparseable shortcuts
    if unparseable_shortcuts:
        print(f"Warning: {len(unparseable_shortcuts)} shortcuts cannot be parsed: {', '.join(unparseable_shortcuts)}")
        if not WIN32_AVAILABLE:
            print("  Install pywin32 package to enable .lnk file parsing")
    
    print("-" * 50)
    
    # Clear launched files at start of each run
    launched_shortcuts.clear()
    
    for filename in all_files:
        basename = os.path.splitext(filename)[0]
        file_type = "native" if is_native_executable(filename) else "shortcut"
        print(f"\nProcessing ({file_type}): {basename}")
        
        # For shortcuts, check if we can parse them
        if filename.endswith('.lnk'):
            targetname, _, _ = get_target_info(filename, args)
            if targetname is None:
                print(f"  ✗ Skipping: Cannot parse shortcut file")
                if not WIN32_AVAILABLE:
                    print(f"     Install pywin32 to enable .lnk file support")
                continue
        
        if not IsFileAlreadyRunning(filename, args):
            print(f"  Launching: {basename}")
            
            if launch_file(filename, args):
                # Mark as launched
                launched_shortcuts.add(filename.lower())
                
                # Wait for process to start
                counter = 0
                targetname, _, _ = get_target_info(filename, args)
                
                # Skip wait if we couldn't determine target
                if targetname is None:
                    print(f"  ! Cannot verify if {basename} started (unable to parse)")
                    time.sleep(args.delay)
                    continue
                
                while counter < args.wait_time:
                    counter += 1
                    time.sleep(1)
                    
                    # For multiple instance programs, just wait a bit
                    if should_allow_multiple(filename, targetname, args):
                        print(f"  Waiting for {basename}... ({counter}/{args.wait_time})")
                        if counter >= 2:  # Shorter wait for known multi-instance programs
                            print(f"  ✓ {basename} launched (multi-instance program)")
                            break
                    else:
                        # Check if process started for single-instance programs
                        if IsFileAlreadyRunning(filename, args):
                            print(f"  ✓ {basename} started successfully")
                            break
                        print(f"  Waiting for {basename} to start... ({counter}/{args.wait_time})")
                
                if counter >= args.wait_time:
                    print(f"  ! {basename} launched but window not detected (may be starting slowly)")
                
                time.sleep(args.delay)  # Configurable pause between launches
        else:
            print(f"  → Skipping {basename} (already running)")
    
    print("\nDesktop initialization complete!")
    print(f"Launched {len(launched_shortcuts)} new file(s)")

if __name__ == "__main__":
    main()