import datetime
import os
import pyautogui
from pyautogui import *
import time
from time import sleep as ____sleep
import re
import psutil
from ShiningPebbles import *

# october constants

# The previous code block did not execute due to an unknown exception.
# I will attempt the conversion again.

# Since the OCR has already been done and the data has been printed in the output,
# I will manually transcribe the output into a Python list.

fund_codes = ['000051', '000052', '100001', '100004', '100008', '100019',
       '100030', '100038', '100042', '100050', '100056', '100059',
       '100060', '100065', '100066', '100069', '100073', '100077',
       '100078', '100079', '805701', '805702', '805703', '805704',
       'A00001', 'A00002', 'A00003', 'A00004', 'A00005']

fund_codes_nov = ['000038', '000039', '000045', '000051', '000052', '000053',
       '100001', '100004', '100008', '100019', '100030', '100038',
       '100042', '100050', '100056', '100059', '100060', '100065',
       '100066', '100069', '100073', '100077', '100078', '100079',
       '805701', '805702', '805703', '805704', 'A00001', 'A00002',
       'A00003', 'A00004', 'A00005']

fund_codes_october = [
    "000051", "000052", "100001", "100004", "100008", "100019", "100030",
    "100038", "100042", "100050", "100056", "100059", "100060", "100065",
    "100066", "100069", "100073", "100077", "100078", "100079", "805701",
    "805702", "805703", "805704", "A00001", "A00002", "A00003", "A00004", "A00005"
]

fund_codes_np = [
    "000038", '000039', '000045', '000053', '100001', '100019', '100030', '100060', '100066'
]

# 2023-12-03 기준
fund_codes_full = ['000002', '000003', '000004', '000005', '000006', '000009',
       '000010', '000011', '000012', '000014', '000018', '000019',
       '000020', '000021', '000022', '000027', '000028', '000029',
       '000038', '000039', '000040', '000041', '000042', '000043',
       '000044', '000045', '000046', '000048', '000050', '000051',
       '000052', '000053', '000054', '000055', '100001', '100002',
       '100003', '100004', '100005', '100006', '100007', '100008',
       '100009', '10000F', '100010', '100011', '100015', '100019',
       '100020', '100021', '100022', '100023', '100024', '100025',
       '100030', '100031', '100033', '100034', '100035', '100036',
       '100037', '100038', '100039', '100042', '100043', '100044',
       '100048', '10004W', '100050', '100051', '100052', '100056',
       '100057', '100058', '100059', '100060', '100061', '100064',
       '100065', '100066', '100067', '100068', '100069', '100070',
       '100072', '100073', '100076', '100077', '100078', '100079',
       '805701', '805702', '805703', '805704', 'A00001', 'A00002',
       'A00003', 'A00004', 'A00005']

start_date_default = '20210101'
end_date_october = '20231031'
end_date_nov = '20231130'
end_date_jan = '20240131'

# global constants
time_duration = 0.25
time_interval_s = 0.25
time_interval = 0.50
time_interval_l = 1
time_interval_loading = 5
time_cooldown = 10
t_10 = 10
t_15 = 15
t_20 = 20
t_30 = 30
t_45 = 45
t_60 = 60

# coordinate constants
# reference screen size: left-half 

screen_width, screen_height = size()
coord_home_input_menu = Point(x=1038, y=42)
coord_home_menu_name = Point(x=1018, y=77)
coord_bos_home_go = Point(x=1099, y=38)
coord_home_logo = Point(x=100, y=34)

coord_8186_input_start_date = Point(x=137, y=173)
coord_8186_input_end_date = Point(x=319, y=176)
coord_8186_input_fund_code = Point(x=505, y=169)
coord_8186_fund_go = Point(x=540, y=173)
coord_8186_fund_popup_checkbox = Point(x=258, y=604)
coord_8186_fund_popup_confirm = Point(x=620, y=1078)
coord_8186_search_go = Point(x=1170, y=125)
coord_8186_excel_go = Point(x=1240, y=240)

coord_2160_input_fund_code = Point(x=107, y=175)
coord_2160_input_start_date = Point(x=535, y=175)
coord_2160_input_end_date = Point(x=630, y=175)
coord_2160_search_go = coord_8186_search_go
coord_2160_excel_go = Point(x=1152, y=219)
coord_2160_excel_format_popup_button = Point(x=1033, y=367)

coord_2205_input_fund_code = Point(x=113, y=211)
coord_2205_input_input_date = Point(x=114, y=241)
coord_2205_search_go = coord_8186_search_go
coord_2205_excel_go = Point(x=1157, y=304)
coord_2205_excel_format_popup_button =Point(x=1035, y=464)

coord_2305_input_fund_code = Point(x=111, y=205)
coord_2305_input_start_date = Point(x=115, y=235)
coord_2305_input_end_date = Point(x=234, y=237)
coord_2305_search_go = coord_8186_search_go
coord_2305_excel_go = Point(x=1148, y=345)
coord_2305_excel_format_popup_button = Point(x=1034, y=501)

coord_2820_input_fund_code = Point(x=121, y=212)
coord_2820_input_start_date = Point(x=120, y=241)
coord_2820_input_end_date = Point(x=237, y=239)
coord_2820_search_go = coord_8186_search_go
coord_2820_excel_go = Point(x=1154, y=280)
coord_2820_excel_format_popup_button = Point(x=1036, y=434)


# key constants
key_save_as_windows = 'F12'
file_format_select_key = 'c'

def preprocess_2110(df_2110_raw):
    df = df_2110_raw.iloc[1:, :]
    df.columns = df_2110_raw.iloc[0, :].tolist()
    df = df[['펀드명', '펀드', '설정일', '만기일', '펀드구분', '채권편입비', '주식편입비', '자본금']]
    return df


def is_dataset_downloaded(menu_code, fund_code, input_date, save_date_yyyymmdd, save_folder_path):
    regex_file_name_snapshot = f'menu{menu_code}-code{fund_code}-date{input_date}-save{save_date_yyyymmdd}'
    regex_file_name_timeseries = f'menu{menu_code}-code{fund_code}-to{input_date}-save{save_date_yyyymmdd}'

    mapping_regex = {
        '2160': regex_file_name_timeseries,
        '2205': regex_file_name_snapshot,
        '2305': regex_file_name_timeseries,
        '2820': regex_file_name_timeseries,
        '8186': regex_file_name_timeseries,
    }

    regex_file_name = mapping_regex[menu_code]

    print(f"- is dataset regex '{regex_file_name}' at '{save_folder_path}' downloaded?")
    lst = scan_files_including_regex(save_folder_path, regex_file_name)
    if len(lst) ==0:
        print(f'- no: {fund_code} not in {save_folder_path}')
        return False
    else:
        print(f'- yes: {fund_code} in {save_folder_path}')
        return True

def are_datasets_downloaded(menu_code, fund_codes, input_date, save_date_yyyymmdd, save_folder_path):
    dct = {}
    bucket_downloaded = []
    bucket_not_downloaded = []
    for fund_code in fund_codes:
        is_downloaded = is_dataset_downloaded(menu_code, fund_code, input_date, save_date_yyyymmdd, save_folder_path)
        if is_downloaded:
            bucket_downloaded.append(fund_code)
        else:
            bucket_not_downloaded.append(fund_code)
    dct[f'{menu_code}_downloaded'] = bucket_downloaded
    dct[f'{menu_code}_not_downloaded'] = bucket_not_downloaded
    return dct
    
def wait_for_loading_to_disappear(image_path, timeout=300):
    """
    loading_image: 로딩 화면의 스크린샷 파일 경로 (예: 'loading.png')
    timeout: 함수가 지정한 시간(초) 동안 로딩 화면이 사라지지 않으면 중단 (기본값: 5분)
    """
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            print("Timeout! Loading screen did not disappear within the specified time.")
            return False
        
        try:
            # locateOnScreen 함수를 사용하여 로딩 화면이 화면에 있는지 확인
            location = pyautogui.locateOnScreen(image_path, confidence=0.8)
            if location is None:
                return True  # 로딩 화면이 없으면 함수 종료
        except ImageNotFoundException:
            print("Loading image not found on the screen, assuming loading is complete.")
            return True  # 이미지를 찾지 못하는 예외 발생 시, 로딩이 완료된 것으로 가정하고 함수 종료
        
        # 로딩 화면이 화면에 있으면 잠시 대기 후 다시 확인
        time.sleep(1)


def click_image(image_path, confidence_level=0.8, timeout=10):
    """
    image_path: 찾고자 하는 이미지의 파일 이름
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
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence_level)
            if location:
                pyautogui.moveTo(location)
                return True
        except ImageNotFoundException:
            # 이미지를 찾지 못한 경우, 예외 처리로 False 반환
            print(f"Could not find the image '{image_path}'. Assuming it is not present.")
            return False

        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print(f"Timeout! Could not find the image '{image_path}' within {timeout} seconds.")
            return False
        
        time.sleep(1)

def is_image_present(image_path, confidence_level=0.8, timeout=2):
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=confidence_level)
            if location:
                return True
        except ImageNotFoundException:
            # 이미지를 찾지 못한 경우, 예외 처리로 False 반환
            print(f"- Could not find the image '{image_path}'. Assuming it is not present.")
            return False

        # 현재 시간과 시작 시간의 차이가 timeout보다 크면 종료
        if time.time() - start_time > timeout:
            print(f"- Timeout! Could not find the image '{image_path}' within {timeout} seconds.")
            return False
        
        time.sleep(1)


def cooltime(t=time_cooldown):
    time.sleep(t)
    return None

def wait_loading():
    image_folder = './image'
    image_name = 'image-8186-loading.png'
    image_path = os.path.join(image_folder, image_name)
    try:
        if is_image_present(image_path, timeout=1):
            print('- step: wait loading...')
            if wait_for_loading_to_disappear(image_path, timeout=300):
                return True
            else:
                print('Error: loading screen did not disappear within the specified time.')
                return False
        else:
            print('Loading image is not present on the screen.')
            return True
    except ImageNotFoundException:
        print('Loading image not found on the screen, continuing with the next steps.')
        return True


def wait_execution():
    image_folder = './image'
    image_name = 'image-excel-header.png'
    image_path = os.path.join(image_folder, image_name)

    if is_image_present(image_path, timeout=30):
        return True
    else:
        return False

def control_from_home_to_menu(menu_code):
    print(f'- step: open menu: {menu_code}')
    click(coord_home_input_menu)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(menu_code))
    time.sleep(time_interval_l)
    press('enter')
    time.sleep(time_interval_l)
    # click(coord_home_menu_name)
    # time.sleep(time_interval)
    # click(coord_bos_home_go)
    # time.sleep(time_interval)

# 8186
def control_on_8186_to_fund_popup(start_date, end_date, fund_code):
    print(f'- step: open fund popup.')
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
    # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
        # 로딩 화면이 사라지면 특정 좌표를 클릭
        if wait_for_loading_to_disappear('image-excel-load-message.png'):
            print(f'- step: execute Excel.')
            click(coord_8186_excel_go)
            cooltime(10)
            if is_image_present('excel-caution-popup.png', timeout=5):
                press('enter')
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
    cooltime()
    if is_image_present('image-excel-header.png', timeout=30):
        os.system("taskkill /im EXCEL.EXE /f")
        print(f'- step: return to home.')

def terminate_excel():
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == 'EXCEL.EXE':
            process.terminate()
            process.wait()
            print("- step: terminate Microsoft Excel.")
            return
    print("- Microsoft Excel is not in process.")

def close_excel():
    print(f'- step: terminate Excel.')
    # click_image('image-excel-header.png', timeout=30)
    cooltime()
    if is_image_present('image-excel-header.png', timeout=20):
        os.system("taskkill /im EXCEL.EXE /f")
        print(f'- step: return to home.')

def goto_home():
    cooltime()
    print(f'- step: return to home.')

# 2160
def control_on_2160_to_set_inputs(fund_code, start_date, end_date):
    print(f'- step: set inputs.')
    click(coord_2160_input_fund_code)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(fund_code))

    time.sleep(time_interval)
    click(coord_2160_input_start_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(start_date))
    
    time.sleep(time_interval)
    click(coord_2160_input_end_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(end_date))

    time.sleep(time_interval)
    click(coord_2160_search_go)
    time.sleep(time_interval_l)

def control_on_2160_to_excel():
    print(f'- step: click excel download button.')
    time.sleep(time_interval_l)
        # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
    # 로딩 화면이 사라지면 특정 좌표를 클릭
        print(f'- step: execute Excel.')
        click(coord_2160_excel_go)
        time.sleep(time_interval_l)
        click(coord_2160_excel_format_popup_button)
    time.sleep(time_interval_loading)
    time.sleep(10)

# 2205
def control_on_2205_to_set_inputs(fund_code, end_date):
    input_date = end_date
    print(f'- step: set inputs.')
    click(coord_2205_input_fund_code)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(fund_code))

    time.sleep(time_interval)
    click(coord_2205_input_input_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(input_date))

    time.sleep(time_interval)
    click(coord_2205_search_go)
    time.sleep(time_interval_l)

def control_on_2205_to_excel():
    print(f'- step: click excel download button.')
    time.sleep(time_interval_l)
        # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
        # 로딩 화면이 사라지면 특정 좌표를 클릭
        if wait_for_loading_to_disappear('image-excel-load-message.png'):
            print(f'- step: execute Excel.')
            click(coord_2205_excel_go)
            time.sleep(time_interval_l)
            click(coord_2205_excel_format_popup_button)
            cooltime(15)
            if is_image_present('excel-caution-popup.png', timeout=10):
                click_image('excel-caution-confirm.png')
    cooltime(10)

# 2305
def control_on_2305_to_set_inputs(fund_code, start_date, end_date):
    print(f'- step: set inputs.')
    click(coord_2305_input_fund_code)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(fund_code))

    time.sleep(time_interval)
    click(coord_2305_input_start_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(start_date))
    
    time.sleep(time_interval)
    click(coord_2305_input_end_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(end_date))

    time.sleep(time_interval)
    click(coord_2305_search_go)
    time.sleep(time_interval_l)

def control_on_2305_to_excel():
    print(f'- step: click excel download button.')
    time.sleep(time_interval_l)
        # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
    # 로딩 화면이 사라지면 특정 좌표를 클릭
        print(f'- step: execute Excel.')
        click(coord_2305_excel_go)
        time.sleep(time_interval_l)
        click(coord_2305_excel_format_popup_button)
    time.sleep(time_interval_loading)
    time.sleep(10)


# 2820
def control_on_2820_to_set_inputs(fund_code, start_date, end_date):
    print(f'- step: set inputs.')
    click(coord_2820_input_fund_code)
    time.sleep(time_interval)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(fund_code))

    time.sleep(time_interval)
    click(coord_2820_input_start_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(start_date))
    
    time.sleep(time_interval)
    click(coord_2820_input_end_date)
    hotkey('ctrl', 'a')
    press('backspace')
    time.sleep(time_interval)
    typewrite(list(end_date))

    time.sleep(time_interval)
    click(coord_2820_search_go)
    time.sleep(time_interval_l)

def control_on_2820_to_excel():
    print(f'- step: click excel download button.')
    time.sleep(time_interval_l)
        # 로딩 화면이 사라질 때까지 기다린다.
    if wait_for_loading_to_disappear('image-8186-loading.png'):
    # 로딩 화면이 사라지면 특정 좌표를 클릭
        print(f'- step: execute Excel.')
        click(coord_2820_excel_go)
        time.sleep(time_interval_l)
        click(coord_2820_excel_format_popup_button)
    time.sleep(time_interval_loading)
    time.sleep(10)


class BOS:
    def __init__(self, fund_code, end_date, start_date=None, file_folder=None):
        self.fund_code = fund_code
        self.start_date = start_date
        self.end_date = end_date
        self.menu_code = None
        self.file_folder = file_folder if file_folder else self.set_file_folder()

    def set_file_folder(self, file_folder=None):
        root_foler = "C:\\Users\\danie\\Documents"
        download_folder = file_folder if file_folder is not None else f'DA-dataset-{get_today(form="yyyymmdd")}'
        file_folder = os.path.join(root_foler, download_folder)
        check_folder_and_create_folder(file_folder)
        self.file_folder = file_folder
        return file_folder

    def goto_menu(self, menu_code):
        self.menu_code = menu_code
        control_from_home_to_menu(self.menu_code)
        wait_loading()
        return self

    def download_dataset(self):
        self.download_dataset_8186(),

    def download_dataset_8186(self):
        # self.file_name = f'menu{self.menu_code}-code{self.fund_code}-save{get_today(form="yyyymmdd")}.csv'
        file_name_timeseries = f'menu{self.menu_code}-code{self.fund_code}-to{self.end_date}-save{get_today(form="yyyymmdd")}.csv'
        self.file_name = file_name_timeseries

        if is_dataset_downloaded(self.menu_code, self.fund_code, input_date=self.end_date, save_date_yyyymmdd=get_today(form='yyyymmdd'), save_folder_path=self.file_folder):
            terminate_excel()
        else:
            control_on_8186_to_fund_popup(self.start_date, self.end_date, self.fund_code)
            wait_loading()
            control_on_fund_popup()
            wait_loading()
            control_on_8186_to_excel()
            wait_execution()
            control_on_excel_to_save_as_popup()
            wait_loading()
            control_on_save_as_popup(self.file_name, self.file_folder)
            wait_loading()
            cooltime()
            self.download_dataset_8186()
            print(f'- save complete: {self.file_name} in {self.file_folder}')
            print(f'- step: return to home.')
            goto_home()


class MOS(BOS):
    def download_dataset(self):
        mapping = {
            '2160': self.download_dataset_2160,
            '2205': self.download_dataset_2205,
            # '2305': self.download_dataset_2305,
            '2820': self.download_dataset_2820
        }
        mapping[self.menu_code]()

    def download_dataset_2160(self):
        # self.file_name = f'menu{self.menu_code}-code{self.fund_code}-save{get_today(form="yyyymmdd")}.csv'
        file_name_timeseries = f'menu{self.menu_code}-code{self.fund_code}-to{self.end_date}-save{get_today(form="yyyymmdd")}.csv'
        self.file_name = file_name_timeseries

        if is_dataset_downloaded(self.menu_code, self.fund_code, input_date=self.end_date, save_date_yyyymmdd=get_today(form='yyyymmdd'), save_folder_path=self.file_folder):
            terminate_excel()
        else:
            control_on_2160_to_set_inputs(self.fund_code, self.start_date, self.end_date)
            wait_loading()
            control_on_2160_to_excel()
            wait_execution()
            control_on_excel_to_save_as_popup()
            wait_loading()
            control_on_save_as_popup(self.file_name, self.file_folder)
            wait_loading()
            cooltime()
            self.download_dataset_2160()
            print(f'- save complete: {self.file_name} in {self.file_folder}')
            print(f'- step: return to home.')
            goto_home()


    def download_dataset_2205(self):
        input_date = self.end_date
        file_name_snapshot = f'menu{self.menu_code}-code{self.fund_code}-date{input_date}-save{get_today(form="yyyymmdd")}.csv'
        self.file_name = file_name_snapshot
        if is_dataset_downloaded(self.menu_code, self.fund_code, input_date, get_today(form='yyyymmdd'), self.file_folder):
            terminate_excel()
        else:
            control_on_2205_to_set_inputs(self.fund_code, self.end_date)
            wait_loading()
            control_on_2205_to_excel()
            wait_execution()
            control_on_excel_to_save_as_popup()
            wait_loading()
            control_on_save_as_popup(self.file_name, self.file_folder)
            wait_loading()
            cooltime()
            self.download_dataset_2205()
            print(f'- save complete: {self.file_name} in {self.file_folder}')
            print(f'- step: return to home.')
            goto_home()

    # def download_dataset_2305(self):
    #     file_name_timeseries = f'menu{self.menu_code}-code{self.fund_code}-to{self.end_date}-save{get_today(form="yyyymmdd")}.csv'
    #     self.file_name = file_name_timeseries
    #     if is_dataset_downloaded(self.menu_code, self.fund_code, input_date=f'{self.start_date}_{self.end_date}', save_date_yyyymmdd=get_today(form='yyyymmdd'), save_folder_path=self.file_folder):
    #         terminate_excel()
    #     else:
    #         control_on_2305_to_set_inputs(self.fund_code, self.start_date, self.end_date)
    #         wait_loading()
    #         control_on_2305_to_excel()
    #         wait_execution()
    #         control_on_excel_to_save_as_popup()
    #         wait_loading()
    #         control_on_save_as_popup(self.file_name, self.file_folder)
    #         wait_loading()
    #         cooltime()
    #         self.download_dataset_2305()
    #         print(f'- save complete: {self.file_name} in {self.file_folder}')
    #         print(f'- step: return to home.')
    #         goto_home()



    def download_dataset_2820(self):
        file_name_timeseries = f'menu{self.menu_code}-code{self.fund_code}-to{self.end_date}-save{get_today(form="yyyymmdd")}.csv'
        self.file_name = file_name_timeseries
        if is_dataset_downloaded(self.menu_code, self.fund_code, self.end_date, get_today(form='yyyymmdd'), self.file_folder):
            print(f'-  complete: {self.file_name} in {self.file_folder}')
            terminate_excel()
        else:
            control_on_2820_to_set_inputs(self.fund_code, self.start_date, self.end_date)
            wait_loading()
            control_on_2820_to_excel()
            wait_execution()
            control_on_excel_to_save_as_popup()
            wait_loading()
            control_on_save_as_popup(self.file_name, self.file_folder)
            wait_loading()
            cooltime()
            self.download_dataset_2820()
            print(f'- save complete: {self.file_name} in {self.file_folder}')
            print(f'- step: return to home.')
            goto_home()

def get_office_system(system_name, fund_code, start_date, end_date):
    mapping = {
        'BOS': BOS,
        'MOS': MOS
    }
    return mapping[system_name](fund_code, start_date=start_date, end_date=end_date)


class DownloaderMacro:
    def __init__(self, fund_code, end_date, start_date='20210101'):
        self.fund_code = fund_code
        self.start_date = start_date
        self.end_date = end_date
        self.bos = get_office_system('BOS', fund_code, self.start_date, self.end_date)
        self.mos = get_office_system('MOS', fund_code, self.start_date, self.end_date)

