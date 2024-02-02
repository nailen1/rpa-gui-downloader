### version
### 2023-12-18-13:55


import pandas as pd
import numpy as np
import os
import time
import calendar
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

import re



def test(func):
    start_time = time.time()  # 함수 실행 전 시간 측정
    result = func  # 함수 실행
    end_time = time.time()  # 함수 실행 후 시간 측정
    duration = end_time - start_time  # 실행 시간 계산
    print(f"Function execution took {duration:.9f} seconds.")  # 소수점 둘째자리까지 출력
    return result  # 함수의 결과 반환


def get_today(form="%Y-%m-%d"):
    mapping = {
        "%Y%m%d": datetime.now().strftime("%Y%m%d"),
        "yyyymmdd": datetime.now().strftime("%Y%m%d"),
        "%Y-%m-%d": datetime.now().strftime("%Y-%m-%d"),
        "yyyy-mm-dd": datetime.now().strftime("%Y-%m-%d"),
        "datetime": datetime.now(),
        "%Y%m%d%H": datetime.now().strftime("%Y%m%d%H"),
        "%Y%m%d%H%M": datetime.now().strftime("%Y%m%d%H%M"),
        "save": datetime.now().strftime("%Y%m%d%H%M"),
    }
    today = mapping[form]
    return today


def get_last_day_of_month(year, month):
    year_input = int(year)
    month_input = int(month)
    # monthrange 함수로 해당 월의 일수를 가져옵니다
    _, last_day = calendar.monthrange(year_input, month_input)
    return f"{year_input}-{month_input:02d}-{last_day}"


def get_weekday(date, language='EN'):
    year = int(date.split('-')[0])
    month = int(date.split('-')[1])
    day = int(date.split('-')[2])

    # weekday 함수는 해당 날짜의 요일을 반환합니다 (0 = 월요일, 6 = 일요일).
    day_index = calendar.weekday(year, month, day)
    mapping = {
        'EN-full': ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        'EN': ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        'KR-full': ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"],
        'KR': ["월", "화", "수", "목", "금", "토", "일"],
    }
    weekday = mapping[language][day_index]
    return weekday


def calculate_prior_date_extended(base_date, days=0, months=0, years=0):
    """
    주어진 날짜로부터 n일/월/년 전의 날짜를 계산합니다.

    :param base_date: 기준 날짜 (문자열 형식: 'YYYY-MM-DD')
    :param kwargs: day, month, year 키워드 인자를 통해 이전으로 계산할 일/월/년 수
    :return: 계산된 이전 날짜 (문자열 형식: 'YYYY-MM-DD')
    """
    date_format = "%Y-%m-%d"
    base_date_obj = datetime.strptime(base_date, date_format)

    # years와 months를 처리하기 위해 relativedelta를 사용
    prior_date_obj = base_date_obj - relativedelta(
        days=days, months=months, years=years
    )
    return prior_date_obj.strftime(date_format)


def scan_files_including_regex(file_folder, regex, option="name"):
    with os.scandir(file_folder) as files:
        lst = [file.name for file in files if re.findall(regex, file.name)]

    mapping = {
        "name": lst,
        "path": [os.path.join(file_folder, file_name) for file_name in lst],
    }
    lst_ordered = sorted(mapping[option])
    return lst_ordered


def format_date(date, form="yyyymmdd"):
    date = date.replace("-", "")
    date_dashed = datetime.strptime(date, "%Y%m%d").strftime("%Y-%m-%d")
    mapping = {"yyyymmdd": date, "yyyy-mm-dd": date_dashed}
    return mapping.get(form, date_dashed)


def save_df_to_file(
    df,
    file_folder,
    file_name_var,
    file_extension=".csv",
    archive=False,
    file_folder_archive="./archive",
):
    def get_today(form="%Y%m%d"):
        return datetime.now().strftime(form)

    try:
        save_time = get_today()
        file_name = f"dataset-{file_name_var}-save{save_time}"+file_extension
        file_path = os.path.join(file_folder, file_name)
        if os.path.exists(file_path) and archive:
            df_archive = pd.read_csv(file_path)
            os.makedirs(file_folder_archive, exist_ok=True)
            archive_file_name = "archive-" + file_name
            archive_file_path = os.path.join(file_folder_archive, archive_file_name)
            df_archive.to_csv(archive_file_path, index=False)
            print(f"Archived: {archive_file_path}")
        df.to_csv(file_path, index=False)
        print(f"- Saved: {file_path}")
    except Exception as e:
        print(f"- Error: {e}")


def check_folder_and_create_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f"- Folder '{folder_name}' created.")
    else:
        print(f"- Folder '{folder_name}' already exists.")
