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

#comm_module.Get_UAC()
#한글 컨트롤 설정 START
constants.pbar.update(0)
app = Application(backend='uia').start(r'C:\Program Files (x86)\Hnc\Office 2020\HncUtils\Studio\HancomStudio.exe')
excel = app.window(title_re='.*Hnc')
constants.pbar.update(30)
time.sleep(4)
def_module.click_after_move3(constants.list21)
#def_module.click_after_move(constants.path_hwp_file_option01);

def_module.click_after_move3(constants.list22)
#def_module.click_after_move(constants.path_hwp_file_option02);
constants.pbar.update(50)
def_module.click_after_move3(constants.list23)
#def_module.click_after_move(constants.path_hwp_file_option03);
constants.pbar.update(60)
time.sleep(1)
gui.hotkey('altleft','f4');
gui.hotkey('altleft','f4');

#한글 자동저장 1분처리
program_path_hnc = r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp.exe"
file_path    = "%s\\files\\test.hwp" % (constants.BASE_DIR)

app = Application().start(r'{} "{}"'.format(program_path_hnc, file_path))

time.sleep(5)
constants.pbar.update(70)

gui.press("altleft")
time.sleep(1)
gui.press("k")
gui.press("u")
time.sleep(2)
gui.hotkey("altleft","p")
def_module.tapkeyPress(5)
gui.press("right")
time.sleep(1)
gui.hotkey("altleft","5")
gui.press("1")
gui.hotkey("altleft","6")
constants.pbar.update(80)
time.sleep(1)
gui.press("5")
gui.hotkey("altleft","d")
gui.press("altleft")

constants.pbar.update(90)
time.sleep(1)
gui.press("f")
gui.press("x")
constants.pbar.finish()


