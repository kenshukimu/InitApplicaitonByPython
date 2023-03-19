import pyautogui as gui
import time

path_root = 'image/'
path_menu_file01 = path_root + 'menu_file01.PNG'
path_menu_file_option01 = path_root + 'menu_file_option01.PNG'
path_menu_file_option02 = path_root + 'menu_file_option02.PNG'


def click_after_move(image_path):
    ins = gui.locateCenterOnScreen(image_path)

    if ins is None:
        return None
    print(ins)
    gui.moveTo(ins)
    gui.doubleClick()


def click_after_move2(image_path1, image_path2):
    ins = gui.locateCenterOnScreen(image_path1)
    print(ins)
    if ins is None:
        ins = gui.locateCenterOnScreen(image_path2)
    if ins is None:
        return None

    print(ins)
    gui.moveTo(ins)
    gui.doubleClick()


# 엑셀/파일
click_after_move(path_menu_file01)
time.sleep(1)
# 엑셀/파일/옵션
click_after_move2(path_menu_file_option01, path_menu_file_option02)
