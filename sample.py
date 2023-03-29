import time
from progressbar import ProgressBar, widgets
import pyautogui as gui

"""
# Create a progress bar object with custom settings
pbar = ProgressBar(
    max_value=100,
    widgets=[widgets.Percentage(), ' ', widgets.Bar()]
)

# Define a function to simulate some work being done
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

for hwp in gui.getAllWindows() :
    print(hwp)
"""
hwp = ''

while not len(hwp) == 1:
    print("Not Loading!!")
    hwp = gui.getWindowsWithTitle("한컴오피스");
    print(len(hwp))

    time.sleep(1)

#print(hwp[0])
"""