"""
fileName : hnc_init.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 한글 초기화 처리
"""
import pyautogui as gui
from pywinauto import Application
import constants
import def_module
import time
import logging_config as log

#comm_module.Get_UAC()
#한글 컨트롤 설정 START
constants.pbar.update(0)
Application(backend='uia').start(r'C:\Program Files (x86)\Hnc\Office 2020\HncUtils\Studio\HancomStudio.exe')
#hwp = app.window(title_re='.*Hnc')

start_time = time.time() 
hwp = ''
while not len(hwp) > 0:
    print("Not Loading!!")
    hwp = gui.getWindowsWithTitle("한컴오피스");
    time.sleep(2)
elapsed_time = time.time() - start_time  # Calculate the elapsed time
log.logger.info(f"한글 오피스 로딩에 걸린 시간 Elapsed time: {elapsed_time:.2f} seconds")

constants.pbar.update(30)
time.sleep(3)
def_module.click_after_move3(constants.list21)
#def_module.click_after_move(constants.path_hwp_file_option01);
time.sleep(3)
def_module.click_after_move3(constants.list22)
#def_module.click_after_move(constants.path_hwp_file_option02);
constants.pbar.update(50)
time.sleep(3)
def_module.click_after_move3(constants.list23)
#def_module.click_after_move(constants.path_hwp_file_option03);
constants.pbar.update(60)
time.sleep(3)
gui.hotkey('altleft','f4');
gui.hotkey('altleft','f4');


#한글 자동저장 1분처리
program_path_hnc = r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp.exe"
file_path    = "%s\\files\\test.hwp" % (constants.BASE_DIR)

Application().start(r'{} "{}"'.format(program_path_hnc, file_path))

start_time = time.time() 
hwp = ''
while not len(hwp) > 0:
    print("Not Loading!!")
    hwp = gui.getWindowsWithTitle("test.hwp");
    time.sleep(2)
elapsed_time = time.time() - start_time  # Calculate the elapsed time
log.logger.info(f"한글 파일 로딩에 걸린 시간 Elapsed time: {elapsed_time:.2f} seconds")

time.sleep(2)
constants.pbar.update(70)

gui.press("altleft")
time.sleep(3)
gui.press("k")
gui.press("u")
time.sleep(3)
gui.hotkey("altleft","p")
def_module.tapkeyPress(5)
gui.press("right")
time.sleep(3)
gui.hotkey("altleft","5")
gui.press("1")
gui.hotkey("altleft","6")
constants.pbar.update(80)
time.sleep(2)
gui.press("5")
gui.hotkey("altleft","d")
gui.press("altleft")

constants.pbar.update(90)
time.sleep(2)
gui.press("f")
gui.press("x")
constants.pbar.finish()