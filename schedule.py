import glob
import os
import re
import uuid
from datetime import datetime

import pandas as pd
import pytz


def find_excel_files(directory):
    # search for all excel files in the directory
    # check if the directory exists
    if not os.path.exists(directory):
        return f"{directory} does not exist."
    # search for all excel files
    xlsx_files = glob.glob(os.path.join(directory, '*.xlsx'))
    xls_files = glob.glob(os.path.join(directory, '*.xls'))
    # 如果文件名以 ~$ 开头，说明是临时文件，不需要处理
    xlsx_files = [file for file in xlsx_files if not os.path.basename(
        file).startswith('~$')]
    xls_files = [file for file in xls_files if not os.path.basename(
        file).startswith('~$')]
    # combine the file lists
    all_excel_files = xlsx_files + xls_files

    return all_excel_files


def extract_times_and_timezone(time_string):
    # 正则表达式匹配模式
    pattern = r"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})\s*([A-Z]+)"

    # 搜索匹配
    match = re.findall(pattern, time_string)
    return match


def contains_date(string):
    # 定义正则表达式模式来匹配 dd/mm/yyyy 格式的日期
    pattern = r'\b\d{2}/\d{2}/\d{4}\b'

    # 搜索匹配
    if re.search(pattern, string):
        return True
    else:
        return False


def convert_to_shanghai_time(time_str, original_tz_str):
    # 时区转换
    # 夏令时的时区缩写为 BST，但是 pytz 模块不支持 BST 时区
    if original_tz_str == 'BST':
        original_tz_str = 'Europe/London'
    # 定义原始时区和目标时区
    original_tz = pytz.timezone(original_tz_str)
    target_tz = pytz.timezone('Asia/Shanghai')

    # 解析时间字符串
    time_format = '%d/%m/%Y %H:%M'
    original_time = datetime.strptime(time_str, time_format)

    # 将时间设置为原始时区
    original_time = original_tz.localize(original_time)

    # 转换到目标时区
    target_time = original_time.astimezone(target_tz)

    return target_time.strftime(time_format)


def create_ics_content(df, callback, *col_name):
    # create the .ics file
    ics_content = []

    # Start of the calendar
    ics_content.append("BEGIN:VCALENDAR\n")
    ics_content.append("VERSION:2.0\n")
    ics_content.append("PRODID:-//CQF Schedule Calendar//EN\n")

    ics_content = callback(df, ics_content, *col_name)
    # End of the calendar
    ics_content.append("END:VCALENDAR\n")

    return "".join(ics_content)


def create_ics_english_content(df, ics_content, col_name):
    for index, row in df.iterrows():
        event_date = row[col_name]
        if pd.notna(row['Title']) and pd.notna(row[col_name]) and contains_date(event_date):
            # Extract the event times and timezone
            event_times = extract_times_and_timezone(row['Time'])
            for event_time in event_times:
                ics_content.append("BEGIN:VEVENT\n")
                ics_content.append("UID:" + str(uuid.uuid4().int) + "\n")
                if not event_time:
                    continue
                start_time = convert_to_shanghai_time(
                    f'{event_date} {event_time[0]}', event_time[2])
                end_time = convert_to_shanghai_time(
                    f'{event_date} {event_time[1]}', event_time[2])
                # Convert string to datetime
                start_date_str = datetime.strptime(
                    f'{start_time}', '%d/%m/%Y %H:%M').strftime('%Y%m%dT%H%M%S')
                end_date_str = datetime.strptime(
                    f'{end_time}', '%d/%m/%Y %H:%M').strftime('%Y%m%dT%H%M%S')

                # Event start and end times
                ics_content.append(f"DTSTART:{start_date_str}\n")
                ics_content.append(f"DTEND:{end_date_str}\n")

                # Event title
                ics_content.append(f"SUMMARY:{row['Title']}\n")

                # Event description
                ics_content.append(
                    f"DESCRIPTION:Module{row['Module']}; {row['Type']}\n")

                # End of an event
                ics_content.append("END:VEVENT\n")

    return ics_content


def create_ics_chinese_content(df, ics_content, col_name):
    for index, row in df.iterrows():
        if pd.notna(row['Title']) and pd.notna(row[col_name]):
            # Start of an event
            ics_content.append("BEGIN:VEVENT\n")
            ics_content.append("UID:" + str(uuid.uuid4().int) + "\n")

            # Formatting the date
            event_date = row[col_name]
            # Convert string to datetime
            start_date_str = event_date.strftime("%Y%m%dT190000")
            end_date_str = event_date.strftime("%Y%m%dT220000")

            # Event start and end times
            ics_content.append(f"DTSTART:{start_date_str}\n")
            ics_content.append(f"DTEND:{end_date_str}\n")

            # Event title
            ics_content.append(f"SUMMARY:{row['Title']}\n")

            # Event description
            ics_content.append(
                f"DESCRIPTION:Module{row['Module']}; {row['Type']}; {row['Golden Tutor']}\n")

            # End of an event
            ics_content.append("END:VEVENT\n")

    return ics_content


def fill_merged_cells(df):
    # 使用前向填充 (ffill) 来填充 NaN 值
    # axis=0 表示沿着列的方向填充（即填充同一列中的 NaN）
    return df.fillna(method='ffill', axis=0)


def write_ics_file(file_path, ics_content):
    # wirte the .ics file
    with open(file_path, 'w') as f:
        f.write(ics_content)


def main():
    directory_path = 'source'
    excel_files = find_excel_files(directory_path)
    output_directory = 'ics'

    for excel_file in excel_files:
        # get the file name
        file_name = os.path.basename(excel_file).split('.')[0]
        df = pd.read_excel(excel_file)
        df_filled = fill_merged_cells(df)

        # Create the contents for both .ics files
        ics_english_content = create_ics_content(
            df_filled, create_ics_english_content, 'Date')
        ics_chinese_content = create_ics_content(
            df_filled, create_ics_chinese_content, 'Chinese date\n(7-10pm)')

        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        # Save the contents to files
        ics_file_path = f'{output_directory}/{file_name}_english.ics'
        ics_chinese_date_file_path = f'{output_directory}/{file_name}_chinese.ics'
        write_ics_file(ics_file_path, ics_english_content)
        write_ics_file(ics_chinese_date_file_path, ics_chinese_content)


if __name__ == "__main__":
    main()
