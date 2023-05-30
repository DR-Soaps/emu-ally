import os
import sys

def install_pywin32():
    try:
        import pip
    except ImportError:
        print("pip is not installed. Please install pip manually.")
        return False

    try:
        import win32api
        print("pywin32 is already installed.")
        return True
    except ImportError:
        print("pywin32 is not found. Installing pywin32...")
        try:
            pip.main(['install', 'pywin32'])
            print("pywin32 has been successfully installed.")
            return True
        except Exception as e:
            print("Failed to install pywin32:", str(e))
            return False

def list_available_drives():
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    return drives

def make_directories(drive):
    base_path = os.path.join(drive, 'Roms')

    if not os.path.exists(base_path):
        os.makedirs(base_path)
    
    platforms = ['bios', 'nes', 'snes', 'n64', 'psp', 'wiiu', 'ps3', 'gba', 'gb', 'gbc', 'gamegear', 'mastersystem',
                 'gen', 'gamecube', 'dreamcast', 'arcade', 'neogeo', 'atari', 'segacd', 'sega32x', 'pcengine',
                 '3ds', 'ps2', 'ps1', 'wii', 'nintendoswitch']
    
    for platform in platforms:
        platform_path = os.path.join(base_path, platform)
        if not os.path.exists(platform_path):
            os.makedirs(platform_path)

    print(f"Directories created on {drive} drive. Look for {base_path}.")

# Check if pywin32 is installed and install it if needed
if not install_pywin32():
    print("pywin32 installation failed. Exiting...")
    sys.exit(1)

import win32api

drives = list_available_drives()
if len(drives) > 0:
    print("Available drives to make emu files on:")
    for i, drive in enumerate(drives):
        print(f"{i+1}. {drive}")

    drive_choice = input("Type the number of the drive to create directories: ")
    try:
        drive_choice = int(drive_choice)
        if 1 <= drive_choice <= len(drives):
            selected_drive = drives[drive_choice - 1]
            make_directories(selected_drive)
        else:
            print("Invalid drive choice.")
    except ValueError:
        print("Invalid drive choice.")
else:
    print("No drives found.")

input("Press Enter to exit.")
