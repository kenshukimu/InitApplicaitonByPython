import time
from progressbar import ProgressBar, widgets
import pyautogui as gui

from win32 import win32api, win32gui
from operator import itemgetter

import screeninfo

from PIL import ImageGrab
from functools import partial

import def_module

"""
# Create a progress bar object with custom settings
pbar = ProgressBar(
    max_value=100,
    widgets=[widgets.Percentage(), ' ', widgets.Bar()]
)

#a
#  Define a function to simulate some work being done
def do_work():
    for i in range(100):
        # Simulate some work being done
        time.sleep(0.1)
        
        # Update the progress bar manually
        pbar.update(i+1)

# Call the function
do_work()

# Finalize the progress bar
pbar.finish()
"""

from progressbar import ProgressBar, widgets

custom_widgets = [
    widgets.Percentage(),
    ' ',
    widgets.Bar(marker='█', left='[', right=']'),
    ' ',
    widgets.Timer(),
    ' ',
    widgets.ETA(),
    ' ',
    widgets.FileTransferSpeed()
]

# Create a progress bar object with custom settings
pbar = ProgressBar(
    max_value=100,
    widgets=custom_widgets,
    #prefix='진행률: ',
    #suffix=' (완료)',
    redirect_stdout=True
)

def do_work():
    for i in pbar(range(100)):
        time.sleep(0.01)

def return_test(itemList) :
    rtn = False
    for item in itemList :
        if item == 1 :
            return True
        else:
            print(item)
            #return False
    return rtn
class MoveMonitor():
    def __init__(self, parent = None):
        monitor = win32api.EnumDisplayMonitors()
        #self.monitorMap = list()
        self.monitorMap = []

        #if(len(monitor) == 1) :
        #    raise Exception("모니터가 한대만 존재합니다.")  

        for info in monitor:
            # 주 모니터와 서브 모니터 구분
            if info[2][0] == 0 and info[2][1] == 0:
                monitorType = "P"
            else :
                monitorType = "S"
            
            self.monitorMap.append({'type': monitorType, 'handle' : info[0], 'left' : info[2][0], 'top' : info[2][1]})    

            print("INFO : "  + str(info))    

    def getActiveWindowHandle(self, titleName):
        def callback(hwnd, hwnd_list: list):
            #activeTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
           
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

        #if(len(hwnd) == 0) :
        #    raise Exception("지정한 Application이 실행되어 있지 않습니다.")  

        # 선택된 프로그램의 현재 모니터 위치
        thisMonitorHandle = win32api.MonitorFromWindow(hwnd[0][1])

        monitor_info = win32api.GetMonitorInfo(thisMonitorHandle)
        print("MONITOR_INFO : " + str(monitor_info))
        print("MONITORMAP : " + str(self.monitorMap))
        #print(str(self.monitorMap[0]['left']))
        print("SCREEN_INFO : " + str(screeninfo.get_monitors()));

        app_x, app_y = 0, 0
        for m in screeninfo.get_monitors():
            if m.width_mm < m.height_mm:
                app_x = m.x + m.width // 2
                app_y = m.y + m.height // 2

                print(str(app_x));
                print(str(app_y));
                break

        # 창 이동
        win32gui.MoveWindow(hwnd[0][1], app_x, app_y, hwnd[0][4], hwnd[0][5], True)
        #if(selMonitor == 0) :
            #win32gui.MoveWindow(hwnd[0][1], self.monitorMap[1]['left'], self.monitorMap[1]['top'], hwnd[0][4], hwnd[0][5], True)
        #elif (selMonitor == 1) :
            #win32gui.MoveWindow(hwnd[0][1], self.monitorMap[0]['left'], self.monitorMap[0]['top'], hwnd[0][4], hwnd[0][5], True)

class MoveMonitor1() :
    def getWindowList(self):
        def callback(hwnd, hwnd_list: list):
            title = win32gui.GetWindowText(hwnd)
            if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
                hwnd_list.append((title, hwnd))
            return True
        output = []
        win32gui.EnumWindows(callback, output)
        return output

def ScreenInfo() :
    app_x, app_y = 0, 0
    for m in screeninfo.get_monitors():
        """
        if m.x == 0 and m.y == 0:
            app_x = (m.width - app_width) // 2
            app_y = (m.height -app_height) // 2
            break
        """
        print(str(m))
"""
if __name__ == "__main__":
    monitor = MoveMonitor1()
    print("\n".join("{: 9d} {}".format(h, t) for t, h in monitor.getWindowList()))
"""
    
#itemlist = [0,2,3]
#print(return_test(itemlist))

#rtn = def_module.click_after_move3(constants.list01)
#print(rtn)

# Call the function
#do_work()
#print(pbar.widgets[0])
#print(pbar.widgets[1])
#print(pbar.widgets[2])
#print(pbar.widgets[3])
#print(pbar.widgets[4])
#print(pbar.widgets[5])

#fwAll = gui.getAllWindows()
#for w in fwAll :
#    print(w)


#import psutil

#for process in psutil.process_iter():
#    print(process.name() + "\t"+str(process.pid))

#for hwp in gui.getAllWindows() :
#    print(hwp)
"""
hwp = ''

while not len(hwp) == 1:
    print("Not Loading!!")
    hwp = gui.getWindowsWithTitle("한컴오피스");
    print(len(hwp))

    time.sleep(1)

#print(hwp[0])
"""
import constants
try :
    #test = MoveMonitor()
    #test.moveWindow("test.xlsx - Excel", 0)
    #test.moveWindow("마우스 속성", 0)
    #ScreenInfo()
    ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
    #img = ImageGrab.grab()
    #img.save("image1.png")  # 파일로 저장 image1.png
    def_module.click_after_move_stop(constants.list21, 1)
except Exception as e:
        print(str(e))
    