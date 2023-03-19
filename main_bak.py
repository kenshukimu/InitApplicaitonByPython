"""
fileName : main_x86.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : x86 실행파일
"""
import pyautogui as gui
import time
import constants
import def_module

"""
pyinstaller 실행방법

1. pyinstaller -F main.py
2. main.spec에서 
   a = Analysis(
    ['main.py'],
    pathex=['D:\\Project\\Pyautogui'],
    binaries=[],
    datas=[('./files/*', './image/*'))],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
2. pyinstaller -w -F --add-data main.spec
"""

# pyautogui 기본세팅 START
# 모든 동작에 1초씩 sleep을 적용
#gui.PAUSE = 1
# pyautogui 기본세팅 END

#레지스트리 EXCEL INIT처리
import comm_module

#MOUSE 컨트롤 설정 START
gui.hotkey('winLeft','r');
gui.write("main.cpl");
gui.press("enter");

time.sleep(2);
def_module.tapkeyPress(5);
gui.press("right");
gui.press("right");
gui.press("right");
rtn = def_module.click_after_move(constants.path_menu_file01);
                
if rtn is None :
     gui.press("tab");

def_module.tapkeyPress(3);
gui.press("enter");

time.sleep(5);
#MOUSE 컨트롤 설정 END

#EXCEL 컨트롤 설정 
import excel_init

#ACCESS 컨트롤 설정 
import access_init

#레지스트리 세팅 실행
def_module.update_regedit_single_binary()
comm_module.update_regedit()
