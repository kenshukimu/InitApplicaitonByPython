"""
fileName : access_init.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 엑세스 초기화 처리
"""
import pyautogui as gui
from pywinauto import Application
import constants
import def_module
import time
#import comm_module

def_module.kill_process("MSACCESS.EXE");
time.sleep(1)

#program_path = r"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.exe"
#file_path    = "%s\\files\\test.accdb" % (sys.path[0], )

#BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

#program_path = r"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.exe"
#file_path    = "%s\\files\\test.accdb" % (constants.BASE_DIR)

#app = Application().start(r'{} "{}"'.format(program_path, file_path))
constants.pbar.update(0)
program_path_access = constants.program_path + "\\MSACCESS.exe"
file_path    = "%s\\files\\test.accdb" % (constants.BASE_DIR)

app = Application().start(r'{} "{}"'.format(program_path_access, file_path))

time.sleep(3)

w = gui.getWindowsWithTitle('Access')[0]  
#활성화 되지 않았다면
if w.isActive == False :
    exit();

if w.isMaximized == True :
    gui.hotkey('winleft','down');

#파일
#def_module.click_after_move(constants.path_access_file_option01)
#옵션
#def_module.click_after_move(constants.path_access_file_option02)
gui.hotkey('altleft','f');
gui.press("t");
time.sleep(1)
constants.pbar.update(10)
#하드웨어 그래픽 사용안함
def_module.click_after_move2(constants.path_access_file_option03_1,constants.path_access_file_option03)
#객체디자이너
def_module.click_after_move(constants.path_access_file_option04)
gui.press("tab");
constants.pbar.update(20)

image_location = None

while image_location is None:
    image_location = gui.locateCenterOnScreen(constants.path_access_file_option07)
    gui.scroll(-500);
    time.sleep(1)

def_module.click_after_move2(constants.path_access_file_option05_1, constants.path_access_file_option05)
def_module.click_after_move2(constants.path_access_file_option06_1, constants.path_access_file_option06)
time.sleep(3);

constants.pbar.update(30)
#리본사용자지정
def_module.click_after_move(constants.path_ribbon_option01)
def_module.tapkeyPress(9)

gui.press("down");
gui.press("down");

gui.press("space");
time.sleep(1);
gui.press("space");
time.sleep(1);
constants.pbar.update(50)
#빠른 실행 도구모음
def_module.click_after_move(constants.path_fast_option01)
def_module.tapkeyPress(7)

gui.press("down");
gui.press("down");

gui.press("space");
time.sleep(1);
gui.press("space");
time.sleep(1);
constants.pbar.update(70)
#보안센터
def_module.click_after_move(constants.path_secu_option01)
#보안센터 설정
def_module.click_after_move2(constants.path_secu_option02_1, constants.path_secu_option02)
#메크로설정
def_module.click_after_move(constants.path_secu_option03)
#모든 메크로 포함
def_module.click_after_move2(constants.path_secu_option04_1, constants.path_secu_option04)
constants.pbar.update(90)
#확인
def_module.click_after_move2(constants.path_menu_file_option07_1,constants.path_menu_file_option07)
def_module.click_after_move2(constants.path_menu_file_option07_1,constants.path_menu_file_option07)

#def_module.kill_process("MSACCESS.EXE");
gui.hotkey('altleft','f4');
constants.pbar.finish()