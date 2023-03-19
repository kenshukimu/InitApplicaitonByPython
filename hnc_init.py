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
#EXCEL 컨트롤 설정 START
constants.pbar.update(0)
app = Application(backend='uia').start(r'C:\Program Files (x86)\Hnc\Office 2020\HncUtils\Studio\HancomStudio.exe')
excel = app.window(title_re='.*Hnc')
constants.pbar.update(30)
time.sleep(4)
def_module.click_after_move(constants.path_hwp_file_option01);
constants.pbar.update(50)
def_module.click_after_move(constants.path_hwp_file_option02);
constants.pbar.update(70)
def_module.click_after_move(constants.path_hwp_file_option03);
constants.pbar.update(90)
time.sleep(1)
gui.hotkey('altleft','f4');
gui.hotkey('altleft','f4');
constants.pbar.finish()
