"""
fileName : def_module.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 모듈모음
"""
import pyautogui as gui
import psutil
import time
from winreg import *
import logging_config as log
import os

#이미지에 맞는 좌표값 찾아서 클릭
def click_after_move(image_path):
    ins = gui.locateCenterOnScreen(image_path)  
    
    if ins is None:
        #print({image_path, ":" , ins})
        log.logger.warning(image_path + " : 이미지를 찾을 수 없습니다.")
        return None
    
    gui.moveTo(ins);
    gui.click();
    time.sleep(1);
    return ""; #return default none

#이미지(2개)에 맞는 좌표값 찾아서 클릭
def click_after_move2(image_path1, image_path2):
    #해상도 정확도 조정(디폴트값 : 99.9)
    ins = gui.locateCenterOnScreen(image_path1)
    if ins is None:
        #print(image_path1 + " : 이미지를 찾을 수 없습니다.")
        log.logger.warning(image_path1 + " : 이미지를 찾을 수 없습니다.")
        ins = gui.locateCenterOnScreen(image_path2)
    if ins is None:
        #print(image_path2 + " : 이미지를 찾을 수 없습니다.")
        log.logger.warning(image_path2 + " : 이미지를 찾을 수 없습니다.")
        return None                                         

    gui.moveTo(ins);
    gui.click();
    time.sleep(1);
    return "";

#이미지에 맞는 좌표값 찾아서 클릭 (없으면 정지)
def click_after_move_stop(image_path):
    ins = gui.locateCenterOnScreen(image_path)  
    
    if ins is None:
        #print({image_path, ":" , ins})
        log.logger.warning(image_path + " : 이미지를 찾을 수 없습니다.")
        exit()

    gui.moveTo(ins);
    gui.click();
    time.sleep(1);
    return ""; #return default none


#프로세스 제거
def kill_process(process_name):
    for proc in psutil.process_iter():
        #if proc.name() == "EXCEL.EXE":
        if proc.name() == process_name:
            proc.kill()

#지정된 횟수만큼 탭키선택
def tapkeyPress(num):
   for i in range(num):
       gui.press("tab");

# 엑셀 폰트 바이너리 코드 변경(초기화)
def update_regedit_single_binary() :
    try:

        value_data = bytes.fromhex('60 00 00 00 60 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 f4 ff ff \
  ff 00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 81 00 00 00 30 \
  74 ad bc b9 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 06 00 00 00 0c 00 00 00 00 00 00 \
  00 0c 00 00 00 0a 00 00 00 02 00 00 00 00 00 00 00 02 00 00 00 33 00 00 00 \
  00 00 00 00 9f 00 08 40 f4 ff ff ff 00 00 00 00 00 00 00 00 00 00 00 00 e8 \
  03 00 00 00 00 00 81 00 00 00 30 74 ad bc b9 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  07 00 00 00 0d 00 00 00 00 00 00 00 0c 00 00 00 0a 00 00 00 02 00 00 00 00 \
  00 00 00 02 00 00 00 33 00 00 00 00 00 00 00 9f 00 08 40 f1 ff ff ff 00 00 \
  00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 81 00 00 00 30 d1 b9 40 \
  c7 20 00 e0 ac 15 b5 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 07 00 00 00 21 00 00 00 00 00 00 00 14 00 \
  00 00 10 00 00 00 04 00 00 00 05 00 00 00 00 00 00 00 33 00 00 00 00 00 00 \
  00 01 00 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 f1 \
  ff ff ff 00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 81 00 00 \
  00 30 d1 b9 40 c7 20 00 e0 ac 15 b5 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 07 00 00 00 21 00 00 00 00 \
  00 00 00 14 00 00 00 10 00 00 00 04 00 00 00 05 00 00 00 00 00 00 00 33 00 \
  00 00 00 00 00 00 01 00 08 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 f4 ff ff ff \
  00 00 00 00 00 00 00 00 00 00 00 00 90 01 00 00 00 00 00 00 00 40 00 00 54 \
  00 61 00 68 00 6f 00 6d 00 61 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 05 00 00 00 1d 00 00 00 00 00 00 00 \
  0e 00 00 00 0c 00 00 00 02 00 00 00 02 00 00 00 00 00 00 00 33 00 00 00 00 \
  00 00 00 ff 01 01 20 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 f1 ff ff ff 00 00 00 00 00 00 \
  00 00 00 00 00 00 90 01 00 00 00 00 00 81 00 00 00 30 d1 b9 40 c7 20 00 e0 \
  ac 15 b5 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 07 00 00 00 21 00 00 00 00 00 00 00 14 00 00 00 10 00 \
  00 00 04 00 00 00 05 00 00 00 00 00 00 00 33 00 00 00 00 00 00 00 01 00 08 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 \
  00 00 00 00 00 00 00 00 00 00 00 00')
        Registrykey = OpenKey(HKEY_CURRENT_USER, r"SOFTWARE\\Microsoft\\Office\\16.0\\Excel", 0, KEY_SET_VALUE)
    except FileNotFoundError as e:
        log.logger.error(f"레지스트리 바이너리 처리 중 에러가 발생하였습니다.")
        #print(f"Error: {e}")
        pass
    except OSError as e:
        log.logger.error(f"레지스트리 바이너리 처리 중 에러가 발생하였습니다.")
        #print(f"Error: {e}")
    else:    
        SetValueEx(Registrykey, "FontInfoCache", 0,REG_BINARY, value_data)
        print ("Setting <" + "FontInfoCache" + "> with value: " + "HexData")        
        CloseKey(Registrykey)

#레지스트리 삭제
def delete_regedit(sub_path) :

    # Open the key for deletion
    try:
        Registrykey = OpenKey(HKEY_CURRENT_USER, r'{}'.format(sub_path), 0, KEY_ALL_ACCESS)   
    
    except FileNotFoundError as e:
        log.logger.error("[" + sub_path + "] " + str(e))       
        pass
    except OSError as e:
        log.logger.error("[" + sub_path + "] " + str(e))        
    else:
        # Delete the key and close the handle
        DeleteKey(Registrykey, "")
        CloseKey(Registrykey)
        log.logger.info(f"Registry key {sub_path} deleted successfully.")        
        #print(f"Registry key {sub_path} deleted successfully.")


#다량의 레지스트리키 변경
def update_regedit_multi(Key_kb, Field, Sub_Key, value) :  

    parmLen = len(Sub_Key)
    z = 0  # Loop Counter for list iteration
    try:
        #if not os.path.exists(Key):
        #    Key = CreateKey(HKEY_CURRENT_USER, r'{}'.format(Key_kb))

        Registrykey = OpenKey(HKEY_CURRENT_USER, r'{}'.format(Key_kb), 0, KEY_SET_VALUE)
    except FileNotFoundError as e:
        log.logger.error("[" + Key_kb + "] " + str(e))
        #print(f"Error: {e}")
        pass
    except OSError as e:
        err = e
        log.logger.error("[" + Key_kb + "] " + str(e))
        #print(f"Error: {e}")
    else:     
        while z < parmLen:

            value_data = value[z] 
            
            #16진수 변환처리
            if Field[z] == REG_DWORD :
                value_data = int(value[z], 16)
            elif Field[z] == REG_BINARY :
                 value_data = bytes.fromhex(value[z])
            
            SetValueEx(Registrykey, Sub_Key[z], 0, Field[z], value_data)
            log.logger.info("Setting <" + Sub_Key[z] + "> with value: " + value[z] + "update successful")
            z += 1
        CloseKey(Registrykey)
        log.logger.info("Setting <" + Sub_Key[z] + "> with value: " + value[z] + "update successful")

#시스템 사양 정보       
def office_version_check():
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
    print ('OS_Name : ' + QueryValueEx(Wininfo_Key, 'ProductName')[0])
    print ('OS_Root_Directory : ' + QueryValueEx(Wininfo_Key, 'SystemRoot')[0])
    print ('OS_install_Date : ' + str(QueryValueEx(Wininfo_Key, 'installDate')[0]))
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\ComputerName\\ActiveComputerName")
    print ('ComputerName : ' + QueryValueEx(Wininfo_Key, 'ComputerName')[0])
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet001\\Control\\Windows")
    print ('Last_ShutDown_Time : ' + str(QueryValueEx(Wininfo_Key, 'ShutdownTime')[0]))
    Wininfo_Key = OpenKey(HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Office\16.0\Outlook")
    print ('Bitness : ' + str(QueryValueEx(Wininfo_Key, 'Bitness')[0]))

    _version = str(QueryValueEx(Wininfo_Key, 'Bitness')[0])

    return _version
    