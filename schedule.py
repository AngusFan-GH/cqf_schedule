import glob
import os
from datetime import datetime, timedelta

import pandas as pd


def find_excel_files(directory):
    # search for all excel files in the directory
    # check if the directory exists
    if not os.path.exists(directory):
        return f"{directory} does not exist."

    # search for all excel files
    xlsx_files = glob.glob(os.path.join(directory, '*.xlsx'))
    xls_files = glob.glob(os.path.join(directory, '*.xls'))

    # combine the file lists
    all_excel_files = xlsx_files + xls_files

    return all_excel_files


def create_ics_content(df, column_date):
    # create the .ics file
    """
    Create the content of an .ics file from the dataframe for the specified date column, 
    with corrected date handling.
    """
    ics_content = []

    # Start of the calendar
    ics_content.append("BEGIN:VCALENDAR\n")
    ics_content.append("VERSION:2.0\n")
    ics_content.append("PRODID:-//CQF Schedule Calendar//EN\n")

    for index, row in df.iterrows():
        if pd.notna(row['Title']) and pd.notna(row[column_date]):
            # Start of an event
            ics_content.append("BEGIN:VEVENT\n")

            # Formatting the date
            event_date = row[column_date]
            if isinstance(event_date, datetime):
                start_date_str = event_date.strftime("%Y%m%dT%H%M%S")
                end_date_str = (event_date + timedelta(hours=2)
                                ).strftime("%Y%m%dT%H%M%S")
            else:
                # Convert string to datetime
                start_date_str = datetime.strptime(
                    str(event_date), '%d/%m/%Y').strftime("%Y%m%dT180000")
                end_date_str = (datetime.strptime(
                    str(event_date), '%d/%m/%Y') + timedelta(hours=2)).strftime("%Y%m%dT200000")

            # Event start and end times
            ics_content.append(f"DTSTART:{start_date_str}\n")
            ics_content.append(f"DTEND:{end_date_str}\n")

            # Event title
            ics_content.append(f"SUMMARY:{row['Title']}\n")

            # Event description
            ics_content.append(
                f"DESCRIPTION:{row['Type']}\n")

            # End of an event
            ics_content.append("END:VEVENT\n")

    # End of the calendar
    ics_content.append("END:VCALENDAR\n")

    return "".join(ics_content)


def create_ics_chinese_content(df, column_date):
    # create the .ics file
    """
    Create the content of an .ics file from the dataframe for the specified date column, 
    with corrected date handling.
    """
    ics_content = []

    # Start of the calendar
    ics_content.append("BEGIN:VCALENDAR\n")
    ics_content.append("VERSION:2.0\n")
    ics_content.append("PRODID:-//CQF Schedule Calendar//EN\n")

    for index, row in df.iterrows():
        if pd.notna(row['Title']) and pd.notna(row[column_date]):
            # Start of an event
            ics_content.append("BEGIN:VEVENT\n")

            # Formatting the date
            event_date = row[column_date]
            if isinstance(event_date, datetime):
                start_date_str = event_date.strftime("%Y%m%dT%H%M%S")
                end_date_str = (event_date + timedelta(hours=2)
                                ).strftime("%Y%m%dT%H%M%S")
            else:
                # Convert string to datetime
                start_date_str = datetime.strptime(
                    str(event_date), '%d/%m/%Y').strftime("%Y%m%dT190000")
                end_date_str = (datetime.strptime(
                    str(event_date), '%d/%m/%Y') + timedelta(hours=3)).strftime("%Y%m%dT220000")

            # Event start and end times
            ics_content.append(f"DTSTART:{start_date_str}\n")
            ics_content.append(f"DTEND:{end_date_str}\n")

            # Event title
            ics_content.append(f"SUMMARY:{row['Title']}\n")

            # Event description
            ics_content.append(
                f"DESCRIPTION:{row['Type']}\n")

            # End of an event
            ics_content.append("END:VEVENT\n")

    # End of the calendar
    ics_content.append("END:VCALENDAR\n")

    return "".join(ics_content)


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

        # Create the contents for both .ics files
        ics_content = create_ics_content(df, 'Date')
        ics_chinese_date_content = create_ics_chinese_content(
            df, 'Chinese date\n(7-10pm)')
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        # Save the contents to files
        ics_file_path = f'{output_directory}/{file_name}.ics'
        ics_chinese_date_file_path = f'{output_directory}/{file_name}_Chinese_Date.ics'
        write_ics_file(ics_file_path, ics_content)
        write_ics_file(ics_chinese_date_file_path, ics_chinese_date_content)


if __name__ == "__main__":
    main()
