"""
fileName : costants.py
author   : (kico) kimhyunsoo
date     : 2023.03.01
desc     : 상수 정의
"""
import os,sys 
import def_module

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

#path_root = 'image/'

BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
path_root    = "%s\\image\\" % (BASE_DIR)

path_menu_file01 = path_root + 'mouse_control01.png'
path_menu_file02 = path_root + 'mouse_control02.png'

path_new_file01 = path_root + 'newFIle01.png'
path_new_file02 = path_root + 'newFIle02.png'

path_menu_file_option01 = path_root + 'menu_file_option01.PNG'
path_menu_file_option02 = path_root + 'menu_file_option02.PNG'
path_menu_file_option03 = path_root + 'menu_file_option03.PNG'
path_menu_file_option05 = path_root + 'menu_file_option05.PNG'
path_menu_file_option05_1 = path_root + 'menu_file_option05_1.PNG'
path_menu_file_option06 = path_root + 'menu_file_option06.PNG'
path_menu_file_option07_1 = path_root + 'menu_file_option07_1.PNG'
path_menu_file_option07 = path_root + 'menu_file_option07.PNG'
path_menu_file_option08_1 = path_root + 'menu_file_option08_1.PNG'
path_menu_file_option08 = path_root + 'menu_file_option08.PNG'

path_save_option01 = path_root + 'save_option01.png'

path_hv_option01 = path_root + 'hv_option01.png'
path_hv_option02 = path_root + 'hv_option02.png'
path_hv_option02_1 = path_root + 'hv_option02_1.png'
path_hv_option03 = path_root + 'hv_option03.png'
path_hv_option03_1 = path_root + 'hv_option03_1.png'
path_hv_option04 = path_root + 'hv_option04.png'
path_hv_option05 = path_root + 'hv_option05.png'

path_ribbon_option01 = path_root + 'ribon_option01.png'
path_ribbon_option02 = path_root + 'ribon_option02.png'
path_ribbon_option02_1 = path_root + 'ribon_option02_1.png'
path_ribbon_option03 = path_root + 'ribon_option03.png'
path_ribbon_option03_1 = path_root + 'ribon_option03_1.png'

path_fast_option01 = path_root + 'fast_option01.png'

path_secu_option01 = path_root + 'secu_option01.png'
path_secu_option02 = path_root + 'secu_option02.png'
path_secu_option02_1 = path_root + 'secu_option02_1.png'
path_secu_option03 = path_root + 'secu_option03.png'
path_secu_option04 = path_root + 'secu_option04.png'
path_secu_option04_1 = path_root + 'secu_option04_1.png'

path_macro_option02 = path_root + 'menu_macro_02.png'

path_hwp_file_option01 = path_root + 'hwp_option01.PNG'
path_hwp_file_option02 = path_root + 'hwp_option02.PNG'
path_hwp_file_option03 = path_root + 'hwp_option03.PNG'

path_access_file_option01 = path_root + 'menu_file_access_option01.PNG'
path_access_file_option02 = path_root + 'menu_file_access_option02.PNG'
path_access_file_option03 = path_root + 'menu_file_access_option03.PNG'
path_access_file_option03_1 = path_root + 'menu_file_access_option03_1.PNG'
path_access_file_option04 = path_root + 'menu_file_access_option04.PNG'
path_access_file_option05 = path_root + 'menu_file_access_option05.PNG'
path_access_file_option05_1 = path_root + 'menu_file_access_option05_1.PNG'
path_access_file_option06 = path_root + 'menu_file_access_option06.PNG'
path_access_file_option06_1 = path_root + 'menu_file_access_option06_1.PNG'
path_access_file_option07 = path_root + 'menu_file_access_option07.PNG'

#python_file_path= os.path.dirname(os.path.abspath(__file__));

officeversion = def_module.office_version_check()
program_path = ''
if officeversion == 'x86' :
    program_path = r"C:\Program Files (x86)\Microsoft Office\Office16"
else :
    program_path = r"C:\Program Files\Microsoft Office\Office16"

