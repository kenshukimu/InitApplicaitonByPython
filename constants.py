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

path_menu_file_option03 = path_root + 'menu_file_option03.PNG'
path_menu_file_option03_1 = path_root + 'menu_file_option03_1.PNG'
#한영자동고침
path_menu_file_option05 = path_root + 'menu_file_option05.PNG'
path_menu_file_option05_1 = path_root + 'menu_file_option05_1.PNG'
path_menu_file_option05_2 = path_root + 'menu_file_option05_2.PNG'
path_menu_file_option05_3 = path_root + 'menu_file_option05_3.PNG'
path_menu_file_option05_4 = path_root + 'menu_file_option05_4.PNG'
path_menu_file_option05_5 = path_root + 'menu_file_option05_5.PNG'

path_menu_file_option05_Not = path_root + 'menu_file_option05_Not.PNG'
path_menu_file_option05_1_Not = path_root + 'menu_file_option05_1_Not.PNG'
path_menu_file_option05_2_Not = path_root + 'menu_file_option05_2_Not.PNG'
path_menu_file_option05_3_Not = path_root + 'menu_file_option05_3_Not.PNG'
path_menu_file_option05_4_Not = path_root + 'menu_file_option05_4_Not.PNG'
path_menu_file_option05_5_Not = path_root + 'menu_file_option05_5_Not.PNG'

path_hv_option01 = path_root + 'hv_option01.png'
path_hv_option01_1 = path_root + 'hv_option01_1.png'
path_hv_option01_2 = path_root + 'hv_option01_2.png'

path_hv_option02 = path_root + 'hv_option02.png'
path_hv_option02_1 = path_root + 'hv_option02_1.png'
path_hv_option02_2 = path_root + 'hv_option02_2.png'
path_hv_option02_3 = path_root + 'hv_option02_3.png'
path_hv_option02_4 = path_root + 'hv_option02_4.png'
path_hv_option02_5 = path_root + 'hv_option02_5.png'

path_hv_option02_Not = path_root + 'hv_option02_Not.png'
path_hv_option02_1_Not = path_root + 'hv_option02_1_Not.png'
path_hv_option02_2_Not = path_root + 'hv_option02_2_Not.png'
path_hv_option02_3_Not = path_root + 'hv_option02_3_Not.png'
path_hv_option02_4_Not = path_root + 'hv_option02_4_Not.png'

path_hv_option03 = path_root + 'hv_option03.png'
path_hv_option03_1 = path_root + 'hv_option03_1.png'
path_hv_option03_2 = path_root + 'hv_option03_2.png'
path_hv_option03_3 = path_root + 'hv_option03_3.png'

path_hv_option03_Not = path_root + 'hv_option03_Not.png'
path_hv_option03_1_Not = path_root + 'hv_option03_1_Not.png'
path_hv_option03_2_Not = path_root + 'hv_option03_2_Not.png'
path_hv_option03_3_Not = path_root + 'hv_option03_3_Not.png'


path_save_option01 = path_root + 'save_option01.png'
path_save_option01_1 = path_root + 'save_option01_1.png'
path_save_option01_2 = path_root + 'save_option01_2.png'

path_ribon_option01 = path_root + 'ribon_option01.png'
path_ribon_option01_1 = path_root + 'ribon_option01_1.png'
path_ribon_option01_2 = path_root + 'ribon_option01_2.png'

path_fast_option01 = path_root + 'fast_option01.png'
path_fast_option01_1 = path_root + 'fast_option01_1.png'
path_fast_option01_2 = path_root + 'fast_option01_2.png'

path_secu_option01 = path_root + 'secu_option01.png'
path_secu_option01_1 = path_root + 'secu_option01_1.png'
path_secu_option01_2 = path_root + 'secu_option01_2.png'

path_hwp_file_option01 = path_root + 'hwp_option01.PNG'
path_hwp_file_option01_1 = path_root + 'hwp_option01_1.PNG'

path_hwp_file_option02 = path_root + 'hwp_option02.PNG'
path_hwp_file_option03 = path_root + 'hwp_option03.PNG'
path_hwp_file_option03_1 = path_root + 'hwp_option03_1.PNG'
path_hwp_file_option03_2 = path_root + 'hwp_option03_2.PNG'

path_access_file_option03 = path_root + 'menu_file_access_option03.PNG'
path_access_file_option03_1 = path_root + 'menu_file_access_option03_1.PNG'
path_access_file_option03_2 = path_root + 'menu_file_access_option03_2.PNG'
path_access_file_option03_3 = path_root + 'menu_file_access_option03_3.PNG'
path_access_file_option03_4 = path_root + 'menu_file_access_option03_4.PNG'

path_access_file_option04 = path_root + 'menu_file_access_option04.PNG'
path_access_file_option04_1 = path_root + 'menu_file_access_option04_1.PNG'
path_access_file_option04_2 = path_root + 'menu_file_access_option04_2.PNG'

path_access_file_option05 = path_root + 'menu_file_access_option05.PNG'
path_access_file_option05_1 = path_root + 'menu_file_access_option05_1.PNG'
path_access_file_option05_2 = path_root + 'menu_file_access_option05_2.PNG'
path_access_file_option05_3 = path_root + 'menu_file_access_option05_3.PNG'
path_access_file_option05_4 = path_root + 'menu_file_access_option05_4.PNG'
path_access_file_option05_5 = path_root + 'menu_file_access_option05_5.PNG'

path_access_file_option06 = path_root + 'menu_file_access_option06.PNG'
path_access_file_option06_1 = path_root + 'menu_file_access_option06_1.PNG'
path_access_file_option06_2 = path_root + 'menu_file_access_option06_2.PNG'
path_access_file_option06_3 = path_root + 'menu_file_access_option06_3.PNG'
path_access_file_option06_4 = path_root + 'menu_file_access_option06_4.PNG'
path_access_file_option06_5 = path_root + 'menu_file_access_option06_5.PNG'

path_access_file_option03_Not = path_root + 'menu_file_access_option03_Not.PNG'
path_access_file_option03_1_Not = path_root + 'menu_file_access_option03_1_Not.PNG'
path_access_file_option03_2_Not = path_root + 'menu_file_access_option03_2_Not.PNG'
path_access_file_option03_3_Not = path_root + 'menu_file_access_option03_3_Not.PNG'

path_access_file_option05_Not = path_root + 'menu_file_access_option05_Not.PNG'
path_access_file_option05_1_Not = path_root + 'menu_file_access_option05_1_Not.PNG'
path_access_file_option05_2_Not = path_root + 'menu_file_access_option05_2_Not.PNG'
path_access_file_option05_3_Not = path_root + 'menu_file_access_option05_3_Not.PNG'

path_access_file_option06_Not = path_root + 'menu_file_access_option06_Not.PNG'
path_access_file_option06_1_Not = path_root + 'menu_file_access_option06_1_Not.PNG'
path_access_file_option06_2_Not = path_root + 'menu_file_access_option06_2_Not.PNG'
path_access_file_option06_3_Not = path_root + 'menu_file_access_option06_3_Not.PNG'

#python_file_path= os.path.dirname(os.path.abspath(__file__));

officeversion = def_module.office_version_check()
program_path = ''
if officeversion == 'x86' :
    program_path = r"C:\Program Files (x86)\Microsoft Office\Office16"
else :
    program_path = r"C:\Program Files\Microsoft Office\Office16"

#엑셀셀
#한영자동고침 이미지
list01 = [path_menu_file_option05_5, path_menu_file_option05_4, path_menu_file_option05_3, path_menu_file_option05_2, path_menu_file_option05_1, path_menu_file_option05]
#확인버튼
#저장
list03 = [path_save_option01, path_save_option01_1, path_save_option01_2]
#고급
list04 = [path_hv_option01, path_hv_option01_1, path_hv_option01_2]
#표시 - 하드웨어 그래픽 사용안함
list06 = [path_hv_option02_5, path_hv_option02_4, path_hv_option02_3, path_hv_option02_2, path_hv_option02_1, path_hv_option02]
#이 워크시트 표시옵션 - 계산 결과 대신 수식줄 셀에 표시
list08 = [path_hv_option03_3, path_hv_option03_2, path_hv_option03_1, path_hv_option03]
#리본 사용자 지정
list09 = [path_ribon_option01, path_ribon_option01_1, path_ribon_option01_2]
#빠른 실행 도구 모임
list10 = [path_fast_option01, path_fast_option01_1, path_fast_option01_2]
#보안센터
list11 = [path_secu_option01, path_secu_option01_1, path_secu_option01_2]
#보안센터 설정
#메크로설정
#모든 메크로 포함
#확인

#엑세스
#하드웨어 그래픽 사용안함
list16 = [path_access_file_option03_4, path_access_file_option03_3, path_access_file_option03_2, path_access_file_option03_1, path_access_file_option03]
#객체디자이너 - 폼 및 보고서 디자인 보기에서 오류검사
list17 = [path_access_file_option04_2, path_access_file_option04_1, path_access_file_option04]
#연결되지 않은 레이블 및 컨트롤 검사
list19 = [path_access_file_option05_5, path_access_file_option05_4, path_access_file_option05_3, path_access_file_option05_2, path_access_file_option05_1, path_access_file_option05]
#연결되지 않은 새 레이블 검사
list20 = [path_access_file_option06_5, path_access_file_option06_4, path_access_file_option06_3, path_access_file_option06_2, path_access_file_option06_1, path_access_file_option06]

#한글
list21 = [path_hwp_file_option01, path_hwp_file_option01_1]
list22 = [path_hwp_file_option02]
list23 = [path_hwp_file_option03_2, path_hwp_file_option03_1, path_hwp_file_option03]

#이미지 체크
#한영자동고침
list24 = [path_menu_file_option05_Not, path_menu_file_option05_1_Not, path_menu_file_option05_2_Not, path_menu_file_option05_3_Not, path_menu_file_option05_4_Not, path_menu_file_option05_5_Not]
#하드웨어 가속화 사용안함(엑셀)
list25 = [path_hv_option02_Not, path_hv_option02_1_Not, path_hv_option02_2_Not, path_hv_option02_3_Not, path_hv_option02_4_Not]
#계산식
list26 = [path_hv_option03_Not, path_hv_option03_1_Not, path_hv_option03_2_Not, path_hv_option03_3_Not]

#연결되지 않은 레이블 및 컨트롤 검사
list27 = [path_access_file_option05_Not, path_access_file_option05_1_Not, path_access_file_option05_2_Not, path_access_file_option05_3_Not]
#연결되지 않은 새 레이블 검사
list28 = [path_access_file_option06_Not, path_access_file_option06_1_Not, path_access_file_option06_2_Not, path_access_file_option06_3_Not]
#하드웨어 가속화 사용안함(엑세스)
list29 = [path_access_file_option03_Not, path_access_file_option03_1_Not, path_access_file_option03_2_Not,  path_access_file_option03_3_Not]