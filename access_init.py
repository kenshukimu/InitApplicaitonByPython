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
import logging_config as log
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

Application().start(r'{} "{}"'.format(program_path_access, file_path))

time.sleep(2)

start_time = time.time() 
access = ''
while not len(access) > 0:
    print("Not Loading!!")
    access = gui.getWindowsWithTitle("Access");
    time.sleep(2)

elapsed_time = time.time() - start_time  # Calculate the elapsed time
log.logger.info(f"엑세스 파일 로딩에 걸린 시간 Elapsed time: {elapsed_time:.2f} seconds")

#w = gui.getWindowsWithTitle('Access')[0]  
#time.sleep(1)
#활성화 되지 않았다면
#while w.isActive == False: 
#    time.sleep(1)

#if w.isMaximized == True :
#    gui.hotkey('winleft','down');

#파일
#def_module.click_after_move(constants.path_access_file_option01)
#옵션
#def_module.click_after_move(constants.path_access_file_option02)
gui.hotkey('altleft','f');
gui.press("t");
time.sleep(1)
constants.pbar.update(10)
#하드웨어 그래픽 사용안함
#def_module.click_after_move2(constants.path_access_file_option03_1,constants.path_access_file_option03)
#def_module.click_after_move3(constants.list16)
rtn = def_module.click_after_move3(constants.list16)

if not rtn :
    rtn = def_module.image_find(constants.list29)
    if not rtn:
        log.logger.warning("하드웨어 그래픽 사용안함을 체크할 수 없습니다.")
    else :
        log.logger.warning("하드웨어 그래픽 사용안함을 체크했습니다.")

#객체디자이너
#def_module.click_after_move(constants.path_access_file_option04)
def_module.click_after_move_stop(constants.list17)
gui.press("tab");
constants.pbar.update(20)

image_location = None

#while image_location is None:
#    image_location = gui.locateCenterOnScreen(constants.path_access_file_option07)
#    gui.scroll(-500);
#    time.sleep(1)

#image_location = None

#while image_location is None:
#    for logImgpath in constants.list18: 
#        n = np.fromfile(logImgpath, np.uint8)
#        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
#        image_location = gui.locateCenterOnScreen(image_path1)  
#        gui.scroll(-500);
#        time.sleep(1)
#        if image_location is not None :
#            break
gui.hotkey("altleft","m")
gui.hotkey("altleft","m")


#연결되지 않은 새 레이블 검사
#def_module.click_after_move3(constants.list20)
rtn = def_module.click_after_move3(constants.list20)

if not rtn :
    rtn = def_module.image_find(constants.list28)
    if not rtn:
        log.logger.warning("연결되지 않은 새 레이블 검사를 체크할 수 없습니다.")
    else :
        log.logger.warning("연결되지 않은 새 레이블 검사를 체크했습니다.")

#def_module.click_after_move2(constants.path_access_file_option05_1, constants.path_access_file_option05)
#def_module.click_after_move2(constants.path_access_file_option06_1, constants.path_access_file_option06)
#연결되지 않은 레이블 및 컨트롤 검사
#def_module.click_after_move3(constants.list19)
rtn = def_module.click_after_move3(constants.list19)

if not rtn :
    rtn = def_module.image_find(constants.list27)
    if not rtn:
        log.logger.warning("연결되지 않은 레이블 및 컨트롤 검사를 체크할 수 없습니다.")
    else :
        log.logger.warning("연결되지 않은 레이블 및 컨트롤 검사를 체크했습니다.")

time.sleep(3);

constants.pbar.update(30)
#리본사용자지정
#def_module.click_after_move(constants.path_ribbon_option01)
def_module.click_after_move3(constants.list09)
def_module.tapkeyPress(9)

gui.press("down");
gui.press("down");

gui.press("space");
time.sleep(1);
gui.press("space");
time.sleep(1);
constants.pbar.update(50)
#빠른 실행 도구모음
def_module.click_after_move3(constants.list10)
def_module.tapkeyPress(7)

gui.press("down");
gui.press("down");

gui.press("space");
time.sleep(1);
gui.press("space");
time.sleep(1);
constants.pbar.update(70)
#보안센터
#def_module.click_after_move(constants.path_secu_option01)
def_module.click_after_move3(constants.list11)
#보안센터 설정
#def_module.click_after_move2(constants.path_secu_option02_1, constants.path_secu_option02)
#def_module.click_after_move3(constants.list12)
gui.hotkey("altleft", "t")
#메크로설정
#def_module.click_after_move(constants.path_secu_option03)
#def_module.click_after_move3(constants.list13)
gui.press("up")
#모든 메크로 포함
#def_module.click_after_move2(constants.path_secu_option04_1, constants.path_secu_option04)
#def_module.click_after_move3(constants.list14)
gui.hotkey("altleft","e")
constants.pbar.update(90)
#확인
def_module.tapkeyPress(2);
gui.press("enter")
def_module.tapkeyPress(1);
gui.press("enter")
#def_module.click_after_move3(constants.list15)
#def_module.click_after_move3(constants.list15)

#def_module.kill_process("MSACCESS.EXE");
gui.hotkey('altleft','f4');
constants.pbar.finish()