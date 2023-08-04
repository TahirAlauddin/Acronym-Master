import winreg

def get_users_desktop_folder():
    # Open the registry key for the desktop folder
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")

    # Get the value of the "Desktop" key
    desktop_path = winreg.QueryValueEx(key, "Desktop")[0]

    return desktop_path
