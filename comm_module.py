"""
fileName : comm_module.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 레지스트리 초기화 처리 (x86)
"""
import os
import sys
from winreg import *
from win32comext.shell import shell
import def_module

#if ctypes.windll.shell32.IsUserAnAdmin(): 
#    print('관리자권한으로 실행된 프로세스입니다.')
#else:
#    print('일반권한으로 실행된 프로세스입니다.')

#권한 변경모듈
def Get_UAC():
    ASADMIN = 'asadmin'

    print(sys.argv)
    if sys.argv[-1] != ASADMIN:
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:] + [ASADMIN])
        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
        sys.exit()

def Windows_info():
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
    print ('OS_Name : ' + QueryValueEx(Wininfo_Key, 'ProductName')[0])
    print ('OS_Root_Directory : ' + QueryValueEx(Wininfo_Key, 'SystemRoot')[0])
    print ('OS_install_Date : ' + str(QueryValueEx(Wininfo_Key, 'installDate')[0]))
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName")
    print ('ComputerName : ' + QueryValueEx(Wininfo_Key, 'ComputerName')[0])
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\Windows")
    print ('Last_ShutDown_Time : ' + str(QueryValueEx(Wininfo_Key, 'ShutdownTime')[0]))

    Wininfo_Key = OpenKey(HKEY_CURRENT_USER,r"Software\Microsoft\Office\16.0\Excel\File MRU\Change")
    print ('ChangeId : ' + str(QueryValueEx(Wininfo_Key, 'ChangeId')[0]))

    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Office\16.0\Outlook")
    print ('Bitness : ' + str(QueryValueEx(Wininfo_Key, 'Bitness')[0]))


global reg_list
reg_list = []
def enum_keys_recursive(key, path):   
    # Enumerate all the subkeys of the given key
    index = 0
    while True:
        try:
            subkey_name = EnumKey(key, index)
            subkey_path = path + "\\" + subkey_name            
            reg_list.append(subkey_path)
            subkey = OpenKey(key, subkey_name)
            enum_keys_recursive(subkey, subkey_path)
            index += 1
        except WindowsError:
            break

def update_regedit() :

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel"
    field_param = [REG_SZ]
    subkey_param = ["ExcelName"]
    value_param = ["Excel"]
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "Software\\Microsoft\\Office\\16.0\\Excel\\File MRU\\Change"
    field_param = [REG_DWORD]
    subkey_param = ["ChangeId"]
    value_param = ["785c1ee7"]
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Options"
    field_param = [REG_DWORD,REG_DWORD,REG_BINARY,REG_SZ,REG_BINARY,REG_DWORD,REG_DWORD,REG_SZ,REG_SZ,REG_DWORD,REG_DWORD,REG_DWORD]
    subkey_param = ["FirstRun","Options5","OptionFormat","Pos","OptionsDlgSizePos","DefaultSheetR2L","UseSystemSeparators","ThousandsSeparator","DecimalSeparator","StickyPtX","StickyPtY","DeveloperTools"]
    value_param = ["00000000","00000080","00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00", "182,182,1440,759", "b6 03 00 00 ad 02 00 00 e5 01 00 00 b2 00 00 00 00 04 00 00", "00000000", "00000001", ",", ".", "000001f4", "000001f5", "00000001"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Place MRU\\Change"
    field_param = [REG_DWORD]
    subkey_param = ["ChangeId"]
    value_param = ["71e54efc"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Recent Templates\\Change"
    field_param = [REG_DWORD]
    subkey_param = ["ChangeId"]
    value_param = ["526f25c5"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Security"
    field_param = [REG_DWORD,REG_DWORD]
    subkey_param = ["VBAWarnings","DataConnectionWarnings"]
    value_param = ["00000001","00000002"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Security\\Trusted Documents"
    field_param = [REG_DWORD]
    subkey_param = ["LastPurgeTime"]
    value_param = ["01a7224c"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

    key_param = "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\StatusBar"
    field_param = [REG_DWORD]
    subkey_param = ["MacroRecord"]
    value_param = ["00000001"];
    def_module.update_regedit_multi(key_param, field_param, subkey_param, value_param)

#Get_UAC()
#Windows_info()
#import platform
#print(platform.architecture())
"""
try:
    key = OpenKey(HKEY_CURRENT_USER, r"SOFTWARE\\Microsoft\\Office\\16.0\\Excel")
    enum_keys_recursive(key, r"SOFTWARE\Microsoft\Office\16.0\Excel")

    for key in reversed(reg_list) :
        def_module.delete_regedit(key)

    def_module.delete_regedit('SOFTWARE\\Microsoft\\Office\\16.0\\Excel')
except FileNotFoundError as e:
    pass

#print(reg_list)
#print(len(reg_list))
try:
    key = OpenKey(HKEY_CURRENT_USER, r"Software\\Microsoft\\VBA\\7.1\\Common")
    enum_keys_recursive(key, r"Software\\Microsoft\\VBA\\7.1\\Common")

    for key in reversed(reg_list) :
        def_module.delete_regedit(key)

    def_module.delete_regedit('Software\\Microsoft\\VBA\\7.1\\Common')
except FileNotFoundError as e:
    pass
"""
