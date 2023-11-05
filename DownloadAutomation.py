import datetime
import os
import pyautogui
from pyautogui import *
import time
from time import sleep as ____sleep

# global constants
time_duration = 0.25
time_interval_s = 0.25
time_interval = 0.50
time_interval_l = 1
time_interval_loading = 5
time_cooldown = 10

# coordinate constants
# reference screen size: left-half 
screen_width, screen_height = size()
coord_bos_home_input_menu = Point(x=1038, y=42)
coord_bos_home_menu_name = Point(x=1018, y=77)
coord_bos_home_go = Point(x=1099, y=38)
coord_8186_input_start_date = Point(x=137, y=173)
coord_8186_input_end_date = Point(x=319, y=176)
coord_8186_input_fund_code = Point(x=505, y=169)
coord_8186_fund_go = Point(x=540, y=173)
coord_8186_fund_popup_checkbox = Point(x=258, y=604)
coord_8186_fund_popup_confirm = Point(x=620, y=1078)
coord_8186_search_go = Point(x=1168, y=126)
coord_8186_excel_go = Point(x=1240, y=240)
coord_home_logo = Point(x=100, y=34)

# key constants
key_save_as_windows = 'F12'
# file_name = f'menu{menu_code}-code{fund_code}-save{get_today()}.csv'
file_format_select_key = 'c'
folder_path = "C:\\rpa-download-datasets"


def get_today():
    return datetime.datetime.now().strftime('%Y%m%d%H%M')

def wait_for_loading_to_disappear(loading_image, timeout=300):
    """
    loading_image: 로딩 화면의 스크린샷 파일 경로 (예: 'loading.png')
    timeout: 함수가 지정한 시간(초) 동안 로딩 화면이 사라지지 않으면 중단 (기본값: 5분)
    """
    start_time = time.time()
    
    while True:
        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print("Timeout! Loading screen did not disappear within the specified time.")
            return False
        
        # locateOnScreen 함수를 사용하여 로딩 화면이 화면에 있는지 확인
        location = pyautogui.locateOnScreen(loading_image, confidence=0.8)
        if location is None:
            return True
        
        # 로딩 화면이 화면에 있으면 잠시 대기 후 다시 확인
        time.sleep(1)

def click_image(image_path, confidence_level=0.8, timeout=10):
    """
    image_path: 찾고자 하는 이미지의 파일 경로
    confidence_level: 이미지 탐색 정확도 (0 ~ 1 사이의 값)
    timeout: 이미지 탐색 최대 대기 시간 (초)
    
    이미지를 찾으면 그 위치를 클릭하고 True를 반환하고, 찾지 못하면 False를 반환한다.
    """
    start_time = time.time()
    while True:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence_level)
        if location:
            pyautogui.click(location)
            return True
        
        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print(f"Timeout! Could not find the image '{image_path}' within {timeout} seconds.")
            return False
        
        time.sleep(1)

def moveTo_image(image_path, confidence_level=0.8, timeout=10):
    """
    image_path: 찾고자 하는 이미지의 파일 경로
    confidence_level: 이미지 탐색 정확도 (0 ~ 1 사이의 값)
    timeout: 이미지 탐색 최대 대기 시간 (초)
    
    이미지를 찾으면 그 위치를 클릭하고 True를 반환하고, 찾지 못하면 False를 반환한다.
    """
    start_time = time.time()
    while True:
        location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence_level)
        if location:
            pyautogui.moveTo(location)
            return True
        
        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print(f"Timeout! Could not find the image '{image_path}' within {timeout} seconds.")
            return False
        
        time.sleep(1)

def is_image_present(image_path, confidence_level=0.8, timeout=2):
    """
    image_path: 찾고자 하는 이미지의 파일 경로
    confidence_level: 이미지 탐색 정확도 (0 ~ 1 사이의 값)
    timeout: 이미지 탐색 최대 대기 시간 (초)
    
    이미지를 찾으면 True를 반환하고, 찾지 못하면 False를 반환한다.
    """
    start_time = time.time()
    while True:
        location = pyautogui.locateOnScreen(image_path, confidence=confidence_level)
        if location:
            return True
        
        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print(f"- Timeout! Could not find the image '{image_path}' within {timeout} seconds.")
            return False
        
        time.sleep(1)

def ____cooltime(t=time_cooldown):
    time.sleep(t)
    return None

def wait_loading():
    if is_image_present('image-8186-loading.png', timeout=1):
        if wait_for_loading_to_disappear('image-8186-loading.png', timeout=300):
            return True
        else:
            print('Error: loading screen did not disappear within the specified time.')
            return False 
    else:
        return True

def wait_execution():
    if is_image_present('image-excel-header.png', timeout=30):
        return True
    else:
        return False

def control_from_home_to_menu(menu_code):
    print(f'- step: open menu: {menu_code}')
    click(coord_bos_home_input_menu)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(menu_code))
    time.sleep(time_interval)
    click(coord_bos_home_menu_name)
    time.sleep(time_interval)
    click(coord_bos_home_go)
    time.sleep(time_interval)

def control_on_8186_to_fund_popup(start_date, end_date, fund_code):
    print(f'- step: oepn fund popup.')
    click(coord_8186_input_start_date)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(start_date))
    time.sleep(time_interval)
    click(coord_8186_input_end_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(end_date))
    time.sleep(time_interval)
    click(coord_8186_input_fund_code)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(fund_code)
    time.sleep(time_interval)
    click(coord_8186_fund_go)
    time.sleep(time_interval_l)

def control_on_fund_popup():
    print(f'- step: select fund on popup.')
    time.sleep(time_interval)
    click(coord_8186_fund_popup_checkbox)
    time.sleep(time_interval)
    click(coord_8186_fund_popup_confirm)

def control_on_8186_to_excel():
    print(f'- step: load system dataset.')
    time.sleep(time_interval)
    click(coord_8186_search_go)
    time.sleep(time_interval)
    # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
        # 로딩 화면이 사라지면 특정 좌표를 클릭
        print(f'- step: execute Excel.')
        click(coord_8186_excel_go)
    time.sleep(time_interval_loading)

def control_on_excel_to_save_as_popup():
    print(f'- step: open save as popup.')
    time.sleep(time_interval_loading)
    press(key_save_as_windows)
    ____sleep(time_cooldown)

def control_on_save_as_popup(file_name, folder_path):
    print(f'- step: input save settings.')
    typewrite(file_name)
    time.sleep(time_interval)
    press('tab')
    ____sleep(time_interval)
    typewrite(file_format_select_key)
    time.sleep(time_interval_loading)
    moveTo_image('image-excel-folder-arrows.png')
    time.sleep(time_interval_l)
    move(screen_width/5,0)
    click()
    time.sleep(time_interval_l)
    typewrite(folder_path)
    time.sleep(time_interval_l)
    hotkey('alt', 's')
    if is_image_present('icon-caution-save-as.png', timeout=1):
        hotkey('alt', 'y')
    time.sleep(time_interval)

def close_excel_and_goto_home():
    print(f'- step: terminate Excel.')
    # click_image('image-excel-header.png', timeout=30)
    ____cooltime()
    if is_image_present('image-excel-header.png', timeout=30):
        hotkey('alt', 'F4')
    click(coord_home_logo)
    print(f'- step: return to home.')


class BOS:

    def __init__(self, fund_code, end_date, start_date='20210101'):
        self.fund_code = fund_code
        self.start_date = start_date
        self.end_date = end_date
        self.menu_code = None

    def goto_menu(self, menu_code):
        self.menu_code = menu_code
        control_from_home_to_menu(self.menu_code)
        wait_loading()
        return self

    def download_dataset(self):
        mapping={
            '8186': self.download_dataset_8186()
        }
        return mapping[self.menu_code]

    def download_dataset_8186(self):
        control_on_8186_to_fund_popup(self.start_date, self.end_date, self.fund_code)
        wait_loading()
        control_on_fund_popup()
        wait_loading()
        control_on_8186_to_excel()
        wait_execution()
        control_on_excel_to_save_as_popup()
        wait_loading()
        self.file_name = f'menu{self.menu_code}-code{self.fund_code}-save{get_today()}.csv'
        self.folder_path = folder_path
        control_on_save_as_popup(self.file_name, self.folder_path)
        wait_loading()
        close_excel_and_goto_home()
        print(f'- save complete: {self.file_name} in {self.folder_path}')

class MOS(BOS):
    def __init__(self, fund_code, end_date, start_date='20210101'):
        super().__init__(fund_code, end_date, start_date)
    
    def download_dataset(self):
        mapping={
            '8186': self.download_dataset_8186()
        }
        return mapping[self.menu_code]



