import winreg
import ctypes
import os


def IsAdmin():
    return ctypes.windll.shell32.IsUserAnAdmin()


def CheckDll():
    return os.path.exists(r"C:\Windows\SysWOW64\clevomof.dll")


def CheckReg():
    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Services\\WmiAcpi", access=winreg.KEY_ALL_ACCESS) as key:
        try:
            value, key_type = winreg.QueryValueEx(key, "MofImagePath")
            if value == r"syswow64\clevomof.dll":
                return True
        except:
            pass
        winreg.SetValueEx(key, "MofImagePath", 0,
                          winreg.REG_EXPAND_SZ, r"syswow64\clevomof.dll")
        return False


if __name__ == '__main__':
    if IsAdmin() == 0:
        print("run as administrator!")
        exit(0)
    if not CheckDll():
        print("No clevomof.dll found! install Control Center and you will get it")
    if CheckReg():
        print("Try to restart")
    else:
        print("Reg fixed!Try to start")
