"""
fileName : excel_init.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 엑셀 초기화 처리
"""
import pyautogui as gui
from pywinauto import Application
import constants
import def_module
import time
import logging_config as log

from shutil import copyfile
from PIL import ImageGrab
from functools import partial

ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)

#import comm_module
copyfile("%s\\files\\Excel.exportedUI" % (constants.BASE_DIR),'C:\\temp\\kcciInitLog\\Excel.exportedUI')
#EXCEL 컨트롤 설정 START
#app = Application(backend='uia').start(r'C:\Program Files (x86)\Microsoft Office\Office16\excel.exe')
#excel = app.window(title_re='.*Excel')

#program_path = r"C:\Program Files (x86)\Microsoft Office\Office16\excel.exe"
#file_path    = "%s\\files\\test.xlsx" % (sys.path[0], )
#app = Application().start(r'{} "{}"'.format(program_path, file_path))
#BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
constants.pbar.update(0)
#program_path = r"C:\Program Files (x86)\Microsoft Office\Office16\excel.exe"
try :
    program_path_excel = constants.program_path + "\\excel.exe"
    file_path    = "%s\\files\\test.xlsx" % (constants.BASE_DIR)
    Application().start(r'{} "{}"'.format(program_path_excel, file_path))

    start_time = time.time() 
    excel = ''

    while not len(excel) > 0:
        print("Not Loading!!")
        excel = gui.getWindowsWithTitle("Excel");
        time.sleep(2)

    elapsed_time = time.time() - start_time  # Calculate the elapsed time
    log.logger.info(f"엑셀 파일 로딩에 걸린 시간 Elapsed time: {elapsed_time:.2f} seconds")

    constants.pbar.update(5)
    # 엑셀/파일
    time.sleep(2)

    #화면이동
    moveClass = def_module.MoveMonitor()
    moveClass.moveWindow("Excel", 0)

    time.sleep(3)

    #rtn = def_module.click_after_move2(constants.path_new_file01, constants.path_new_file02)
    #w = gui.getWindowsWithTitle('Excel')[0]  
    #time.sleep(2)
    #활성화 되지 않았다면
    #if w.isActive == False :
    #    exit();

    #if w.isMaximized == True :
    #    gui.hotkey('winleft','down');

    #print(rtn);
    #if rtn is None :
    #    exit();     
    #파일선택
    #def_module.click_after_move(constants.path_menu_file_option01)
    #옵션선택
    #def_module.click_after_move(constants.path_menu_file_option02)
    gui.hotkey('altleft','f');
    gui.press("t");
    time.sleep(1)

    #언어교정
    #def_module.click_after_move(constants.path_menu_file_option03)
    gui.press("down");
    gui.press("down");

    constants.pbar.update(10)
    #자동고침옵션
    gui.press("tab");
    gui.press("space");
    time.sleep(1);

    gui.press("altleft")

    #한영자동고침(window10, window11)
    #def_module.click_after_move2(constants.path_menu_file_option05_1, constants.path_menu_file_option05)
    rtn = def_module.click_after_move3(constants.list01)

    if not rtn :
        rtn = def_module.image_find(constants.list24)
        if not rtn:
            log.logger.warning("한영자동고침을 체크할 수 없습니다.")
        else :
            log.logger.warning("한영자동고침을 체크했습니다.")

    #def_module.click_after_move2(constants.path_menu_file_option08_1,constants.path_menu_file_option08)
    #def_module.click_after_move3(constants.list02)
    gui.press("tab")
    gui.press("enter")
    constants.pbar.update(20)
    #저장
    #imglist.append(constants.path_menu_file_option08)
    #def_module.click_after_move(constants.path_save_option01)
    #gui.write("1");
    def_module.click_after_move_stop(constants.list03, 1)
    #자동복구간격 1분
    def_module.tapkeyPress(3);
    gui.write("1");
    time.sleep(1);
    constants.pbar.update(30)
    #고급
    #def_module.click_after_move(constants.path_hv_option01)
    def_module.click_after_move_stop(constants.list04)
    gui.press("tab");
    time.sleep(1);
    constants.pbar.update(40)
    #하드웨어 그래픽 가속 사용안함
    #image_location = None

    #while image_location is None:
    #    for logImgpath in constants.list05: 
    #        n = np.fromfile(logImgpath, np.uint8)
    #        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
    #        image_location = gui.locateCenterOnScreen(image_path1)  
    #        gui.scroll(-300);
    #        time.sleep(1)
    #        if image_location is not None :
    #            break
    gui.hotkey("altleft","i")
    gui.hotkey("altleft","i")
    #def_module.click_after_move3(constants.list05)

    #gui.scroll(-3000);
    #def_module.click_after_move2(constants.path_hv_option02_1, constants.path_hv_option02)
    #def_module.click_after_move3(constants.list06)
    rtn = def_module.click_after_move3(constants.list06)

    if not rtn :
        rtn = def_module.image_find(constants.list25)
        if not rtn:
            log.logger.warning("하드웨어 그래픽 가속 사용안함을 체크할 수 없습니다.")
        else :
            log.logger.warning("하드웨어 그래픽 가속 사용안함을 체크했습니다.")

    time.sleep(1);
    constants.pbar.update(50)
    #계산 결과 대신 수식줄 셀에 표시
    #image_location = None

    #while image_location is None:
    #    for logImgpath in constants.list07: 
    #        n = np.fromfile(logImgpath, np.uint8)
    #        image_path1 = cv2.imdecode(n, cv2.IMREAD_COLOR)
    #        image_location = gui.locateCenterOnScreen(image_path1)  
    #        gui.scroll(-300);
    #        time.sleep(1)
    #        if image_location is not None :
    #            break
    gui.hotkey("altleft","z")
    gui.hotkey("altleft","z")
    #def_module.click_after_move3(constants.list07)
    #def_module.click_after_move2(constants.path_hv_option03_1, constants.path_hv_option03)
    #def_module.click_after_move3(constants.list08)
    rtn = def_module.click_after_move3(constants.list08)

    if not rtn :
        rtn = def_module.image_find(constants.list26)
        if not rtn:
            log.logger.warning("계산 결과 대신 수식줄 셀에 표시를 체크할 수 없습니다.")
        else :
            log.logger.warning("계산 결과 대신 수식줄 셀에 표시를 체크했습니다.")

    time.sleep(1);
    constants.pbar.update(60)
    #리본 사용자 지정
    #def_module.click_after_move(constants.path_ribbon_option01)
    #def_module.click_after_move(constants.path_ribbon_option03)
    def_module.click_after_move_stop(constants.list09, 1)

    def_module.tapkeyPress(9)

    gui.press("down");
    gui.press("down");

    gui.press("space");
    time.sleep(1);
    gui.press("space");
    time.sleep(1);
    gui.press("tab");
    gui.press("down");
    gui.press("enter");

    time.sleep(1);

    #file_excel_init_path = "%s\\files\\Excel.exportedUI" % (constants.BASE_DIR)
    file_excel_init_path = 'C:\\temp\\kcciInitLog\\Excel.exportedUI'
    gui.write(file_excel_init_path);
    time.sleep(1);
    gui.press("enter");
    time.sleep(1);
    gui.press("space");
    time.sleep(1);
    constants.pbar.update(70)
    #빠른 실행 도구 모임
    #def_module.click_after_move(constants.path_fast_option01)
    def_module.click_after_move_stop(constants.list10, 1)
    def_module.tapkeyPress(7)

    gui.press("down");
    gui.press("down");

    gui.press("space");
    time.sleep(1);
    gui.press("space");
    time.sleep(1);
    #def_module.click_after_move(constants.path_ribbon_option02)
    gui.press("tab");
    gui.press("down"); 
    gui.press("enter");

    time.sleep(1);

    gui.write(file_excel_init_path);
    time.sleep(1);
    gui.press("enter");
    time.sleep(1);
    gui.press("space");
    time.sleep(1);
    constants.pbar.update(80)
    #보안센터
    #def_module.click_after_move(constants.path_secu_option01)
    def_module.click_after_move_stop(constants.list11, 1)
    #보안센터 설정
    #def_module.click_after_move2(constants.path_secu_option02_1, constants.path_secu_option02)
    #def_module.click_after_move3(constants.list12)
    gui.hotkey("altleft", "t")
    #메크로설정
    #def_module.click_after_move(constants.path_secu_option03)
    #def_module.click_after_move3(constants.list13)
    gui.press("up")
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
    #def_module.click_after_move2(constants.path_menu_file_option07_1,constants.path_menu_file_option07)
    #def_module.click_after_move2(constants.path_menu_file_option07_1,constants.path_menu_file_option07)
    """
    #visualBasic 창처리 ->  HKEY_CURRENT_USER\Software\Microsoft\VBA\7.1\Common 삭제로 해결
    gui.hotkey('altleft','f11');
    win = gui.getActiveWindow()
    if win.isMaximized == False :
        win.maximize()

    time.sleep(1)
    gui.hotkey('altleft','t');
    gui.press('o');

    click_after_move(path_macro_option02)
    gui.press('enter');
    gui.hotkey('altleft','f4');
    """
    moveClass = def_module.MoveMonitor()
    moveClass.moveWindow("Excel", 1)

    time.sleep(1)
    gui.hotkey('altleft','f4');
    constants.pbar.finish()
    #def_module.kill_process("EXCEL.EXE");
    #app.window(title="...").close()
    #EXCEL 컨트롤 설정 END
except Exception as e:
        log.logger.error("엑셀처리에러 : " + str(e))
        print(str(e))
