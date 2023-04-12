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
import numpy as np
import cv2
from win32 import win32api, win32gui
import screeninfo

#이미지 경로 배열로 받아 처리
def click_after_move3(image_path_list, per=None):
    rtn = False

    for logImgpath in image_path_list: 
        n = np.fromfile(logImgpath, np.uint8)
        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
        
        if per is None:
            ins = gui.locateCenterOnScreen(image_path1)
        else :
            ins = gui.locateCenterOnScreen(image_path1)
    
        if ins is None:
            #print(image_path2 + " : 이미지를 찾을 수 없습니다.")
            log.logger.warning(logImgpath + " : 이미지를 찾을 수 없습니다.")
        else:            
            if(monitor_x >= 0) :
                 gui.moveTo(ins.x, ins.y);
            else:
                gui.moveTo(monitor_x + ins.x, ins.y);
            
            gui.click();
            time.sleep(1); 
            rtn = True      
            break     
    return rtn

#이미지에 맞는 좌표값 찾아서 클릭 (없으면 정지)
def click_after_move_stop(image_path_list, per=None):
    rtn = False

    for logImgpath in image_path_list: 

        n = np.fromfile(logImgpath, np.uint8)
        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
        
        if per is None:
            ins = gui.locateCenterOnScreen(image_path1)
        else :
            ins = gui.locateCenterOnScreen(image_path1)    
    
        if ins is None:
            log.logger.warning(logImgpath + " : 이미지를 찾을 수 없습니다.")
        else:   
            #print(monitor_x)
            #print("click_after_move_stop : " + str(monitor_x))
            #print("click_after_move_stop : " + str(ins.x))

            if(monitor_x >= 0) :
                 gui.moveTo(ins.x, monitor_y + ins.y);
            else:
                gui.moveTo(monitor_x + ins.x, monitor_y + ins.y);
           
            gui.click();
            time.sleep(1); 
            rtn = True  
            break

    if not rtn :
        gui.screenshot('C:\\temp\\kcciInitLog\\KcciErrorImg.png')
        raise Exception("![중요]! " + logImgpath + " : 이미지를 찾을 수 없습니다.")       
    
    return rtn

#이미지 확인
def image_find(image_path_list, per=None):
    rtn = False

    for logImgpath in image_path_list: 
        n = np.fromfile(logImgpath, np.uint8)
        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
        
        if per is None:
            ins = gui.locateCenterOnScreen(image_path1)           
        else :
            ins = gui.locateCenterOnScreen(image_path1)
    
        if ins is None:
            #print(image_path2 + " : 이미지를 찾을 수 없습니다.")
            log.logger.warning(logImgpath + " : 이미지를 찾을 수 없습니다.")
        else:
            rtn = True;
            break   
    return rtn


#프로세스 제거
def kill_process(process_name):
    for proc in psutil.process_iter():
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
        #print ("Setting <" + "FontInfoCache" + "> with value: " + "HexData")    
        log.logger.info("Setting <" + "FontInfoCache" + "> with value: " + "HexData")    
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
        Registrykey = OpenKey(HKEY_CURRENT_USER, r'{}'.format(Key_kb), 0, KEY_SET_VALUE)
    except FileNotFoundError as e:
        log.logger.error("[" + Key_kb + "] " + str(e))
        pass
    except OSError as e:
        log.logger.error("[" + Key_kb + "] " + str(e))
        pass
    except Exception as e:
        log.logger.error("[" + Key_kb + "] " + str(e))
        pass
    else:            
        while z < parmLen:

            value_data = value[z] 
            
            #16진수 변환처리
            if Field[z] == REG_DWORD :
                value_data = int(value[z], 16)
            elif Field[z] == REG_BINARY :
                 value_data = bytes.fromhex(value[z])            
          
            SetValueEx(Registrykey, Sub_Key[z], 0, Field[z], value_data)
            z += 1
        log.logger.info("Setting <" + Key_kb + "> update Successed")    
        CloseKey(Registrykey)        

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

#화면 이동 클래스
#Excel, Access, 한글, 한컴오피스, 마우스 속성
class MoveMonitor():
    def __init__(self, parent = None):
        monitor = win32api.EnumDisplayMonitors()
        self.monitorMap = []

        if(len(monitor) == 1) :
            raise Exception("모니터가 한대만 존재합니다.")  

        for info in monitor:
            # 주 모니터와 서브 모니터 구분
            if info[2][0] == 0 and info[2][1] == 0:
                monitorType = "P"
            else :
                monitorType = "S"
            
            self.monitorMap.append({'type': monitorType, 'handle' : info[0], 'left' : info[2][0], 'top' : info[2][1]})        

    def getActiveWindowHandle(self, titleName):
        def callback(hwnd, hwnd_list: list):
            title = win32gui.GetWindowText(hwnd)
            if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
                #hwnd_list.append((title, hwnd))
                if titleName in title:
                    rect = win32gui.GetWindowRect(hwnd)
                    hwnd_list.append((title, hwnd, rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]))
            return True
        output = []
        win32gui.EnumWindows(callback, output)
        return output

    def moveWindow(self, titleName, selMonitor):
        # 선택된 프로그램의 핸들값
        hwnd = self.getActiveWindowHandle(titleName)

        if(len(hwnd) == 0) :
            raise Exception("지정한 Application이 실행되어 있지 않습니다.")  

        # 선택된 프로그램의 현재 모니터 위치
        thisMonitorHandle = win32api.MonitorFromWindow(hwnd[0][1])

        monitor_info = win32api.GetMonitorInfo(thisMonitorHandle)
        #print(str(monitor_info))
        #print(str(self.monitorMap))
        #print(str(self.monitorMap[0]['left']))

        # 창 이동
        app_x, app_y = 0, 0
        for m in screeninfo.get_monitors():
            if m.width_mm < m.height_mm and selMonitor == 0:
                app_x = m.x
                app_y = m.y + m.height // 3

                #이동 모니터 X좌표 가지고 처리
                global monitor_x
                global monitor_y
                
                monitor_x = m.x
                monitor_y = m.y

                break
            elif m.width_mm > m.height_mm and selMonitor == 1:
                app_x = m.x
                app_y = m.y + m.height // 3   
                break             

        # 창 이동
        win32gui.MoveWindow(hwnd[0][1], app_x, app_y, hwnd[0][4], hwnd[0][5], True)

        """
        if(selMonitor == 0) :
            win32gui.MoveWindow(hwnd[0][1], self.monitorMap[1]['left'], self.monitorMap[1]['top'], hwnd[0][4], hwnd[0][5], True)
        elif (selMonitor == 1) :
            win32gui.MoveWindow(hwnd[0][1], self.monitorMap[0]['left'], self.monitorMap[0]['top'], hwnd[0][4], hwnd[0][5], True)
        """