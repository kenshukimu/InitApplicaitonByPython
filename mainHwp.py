"""
fileName : main.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 실행코드
"""
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

#install pyinstaller -w --uac-admin main.py
#pyinstaller -F --uac-admin --add-data "./files/*;./files" --add-data "./image/*;./image" -n kcciOfficeInit main.py

pyinstaller -F --uac-admin --add-data "./files/*;./files" --add-data "./image/*;./image" -n  KcciOfficeInit --icon=./image/file_excel_icon.png main.py 
"""
from winreg import *
import constants
import def_module
import time
import pyautogui as gui
import logging_config as log
from pathlib import Path

try:   
    
    t_start_time = time.time() 
    print("#################################################")
    print("###########대한상공회의소 자격평가사업단###########")
    print("########  엑셀/엑세스/한글 프로그램 초기화  #######")
    print("################################# Ver. 1.0 ######")

    log.logger.info("프로세스 Kill 처리 START")
    def_module.kill_process("Hwp.exe");
    time.sleep(1)
    
    log.logger.info("프로세스 Kill 처리 END")
    
    print("1. 마우스 컨트롤 제어 초기화 시작")
    #MOUSE 컨트롤 설정 START
    constants.pbar.update(10)
    gui.hotkey('winLeft','r');
    constants.pbar.update(20)
    gui.write("main.cpl");
    constants.pbar.update(30)
    gui.press("enter");

    constants.pbar.update(40)
    time.sleep(2);
    def_module.tapkeyPress(5);
    constants.pbar.update(50)
    gui.press("right");
    constants.pbar.update(60)
    gui.press("right");
    constants.pbar.update(70)
    gui.press("right");
    constants.pbar.update(80)
    #rtn = def_module.click_after_move(constants.path_menu_file01);
    #gui.press("tab");
    gui.hotkey("altleft","n")
                    
    #if rtn is None :
    #     gui.press("tab");
    constants.pbar.update(90)
    gui.press("a");
    gui.press("enter");

    time.sleep(1);
    constants.pbar.finish()
    print("   마우스 컨트롤 제어 초기화 종료")
    #MOUSE 컨트롤 설정 END

    start_time = time.time() 
    #한글 초기화 설정 
    print("5. 한글 초기화 시작") 
    import hnc_init
    print("   한글 초기화 종료") 
    elapsed_time = time.time() - start_time  # Calculate the elapsed time
    log.logger.info(f"한글 처리 Elapsed time: {elapsed_time:.2f} seconds")
    
    elapsed_time = time.time() - start_time  # Calculate the elapsed time
    log.logger.info(f"전체 초기화 처리 Elapsed time: {elapsed_time:.2f} seconds")
except Exception as e:
    log.logger.error(f"전체 처리중 에러 발생 : " + str(e))
finally:
    def_module.kill_process("Hwp.exe");