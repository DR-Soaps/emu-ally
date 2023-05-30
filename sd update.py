import os
import shutil
import win32api
import win32con
import win32file
import subprocess

def find_drive_letter(drive_name):
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    for drive in drives:
        if win32file.GetDriveType(drive) == win32con.DRIVE_REMOVABLE:
            if win32api.GetVolumeInformation(drive)[0] == drive_name:
                return drive
    return None

def create_files(drive_letter):
    root_path = os.path.join(drive_letter, 'Roms')
    roms_folders = [
        'bios', 'nes', 'snes', 'n64', 'psp', 'wiiu', 'ps3', 'gba', 'gb', 'gbc',
        'gamegear', 'mastersystem', 'gen', 'gamecube', 'dreamcast', 'arcade',
        'neogeo', 'atari', 'segacd', 'sega32x', 'pcengine', '3ds', 'ps2', 'ps1',
        'wii', 'nintendoswitch'
    ]

    os.makedirs(root_path, exist_ok=True)
    for folder in roms_folders:
        os.makedirs(os.path.join(root_path, folder), exist_ok=True)

def delete_files(drive_letter):
    root_path = os.path.join(drive_letter, 'Roms')
    if os.path.exists(root_path):
        shutil.rmtree(root_path)
        print(f'Deleted files on {drive_letter}')
    else:
        print('Files not found on the selected drive')

def install_package(package_name):
    try:
        subprocess.check_call(['pip', 'install', package_name])
    except subprocess.CalledProcessError:
        print(f'Failed to install {package_name}. Please install it manually.')

def list_drives():
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    for index, drive in enumerate(drives):
        drive_type = win32file.GetDriveType(drive)
        drive_label = win32api.GetVolumeInformation(drive)[0]
        drive_type_str = ''
        if drive_type == win32con.DRIVE_FIXED:
            drive_type_str = 'Local Disk'
        elif drive_type == win32con.DRIVE_REMOVABLE:
            drive_type_str = 'Removable Disk'
        print(f'{index + 1}. {drive} ({drive_label}) - {drive_type_str}')

def main():
    package_name = 'pywin32'
    try:
        import win32api
        import win32con
        import win32file
    except ImportError:
        print(f'{package_name} package not found. Installing...')
        install_package(package_name)
        try:
            import win32api
            import win32con
            import win32file
        except ImportError:
            print(f'Failed to import {package_name}. Please install it manually.')
            return

    print('This script will create or delete files for emulators on the selected drive.')
    print('Available drives:')
    list_drives()

    drive_index = input('Enter the number to select the drive: ')
    try:
        drive_index = int(drive_index)
    except ValueError:
        print('Invalid drive index.')
        return

    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    if drive_index < 1 or drive_index > len(drives):
        print('Invalid drive index.')
        return

    drive_letter = drives[drive_index - 1]
    print(f'Selected drive: {drive_letter}')

    action = input('Enter "create" to create files, "delete" to delete files: ')

    if action == 'create':
        create_files(drive_letter)
        print('Files created successfully.')
    elif action == 'delete':
        delete_files(drive_letter)
    else:
        print('Invalid action.')

if __name__ == '__main__':
    main()
