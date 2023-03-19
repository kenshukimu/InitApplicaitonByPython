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

# Call the function
do_work()
#print(pbar.widgets[0])
#print(pbar.widgets[1])
#print(pbar.widgets[2])
#print(pbar.widgets[3])
print(pbar.widgets[4])
#print(pbar.widgets[5])

#fwAll = gui.getAllWindows()
#for w in fwAll :
#    print(w)
