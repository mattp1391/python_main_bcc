import pandas as pd
import numpy
from tqdm import tqdm
from datetime import datetime, timedelta
import os


input_file = r"C:\General\BCC_Software\Python\python_repository\development_files\analyse_tube_counts\inputs\13759 " \
             r"Eastbound 23 Dobson Street ASCOT Between Racecourse Rd and Seymour Rd Class Volume 15 minute Report " \
             r".txt "

input_folder = r"C:\General\BCC_Software\Python\python_repository\development_files\analyse_tube_counts\inputs"
output_file = r"C:\General\BCC_Software\Python\python_repository\development_files\analyse_tube_counts\outputs" \
              r"\tubes_combined.csv "

def replace_multiple_spaces(text_string):
    corrected_string = " ".join(text_string.split())
    return corrected_string


def find_date_range(text_str):
    date_strings = text_str.split('\t')[1]
    date_strings = date_strings.split(' (')[0]
    date_strings = date_strings.split(' => ')
    return date_strings


def find_site(text_str):
    site = text_str.split('\t')[1]
    site = site.split(' <')[0]
    return site


def find_direction(text_str):
    direction = text_str.split('\t')[1]
    direction = direction.split(' ')[0]
    return direction


def determine_headings(text_str):
    headings = ['site', 'direction', 'date', 'time', 'total']
    class_headings = text_str.split('\t')[1].replace('\n', '').split(',')
    headings = headings + class_headings
    return headings


def find_four_digit_start_time(time_str):
    start_time = time_str.split(' ')[0]
    start_time = start_time.replace(':', '')
    start_time = f"{start_time:04}"
    return start_time


def find_start_date(date_str):
    start_date = datetime.strptime(date_str, '%H:%M %A, %d %B %Y')
    return start_date


def obtain_tube_data(file_name):
    site = None
    direction = None
    date_strings = None
    headings_found = False
    headings = None
    with open(file_name, "r") as f:
        for line in f:
            if site is None:
                if "Site:" in line:
                    site = find_site(line)
            elif date_strings is None:
                if "Filter time:" in line:
                    date_strings = find_date_range(line)
                    start_time = find_four_digit_start_time(date_strings[0])
                    analysis_date = find_start_date(date_strings[0])
            elif headings is None:
                if 'Included classes:' in line:
                    headings = determine_headings(line)
                    df = pd.DataFrame(columns=headings)
            elif direction is None:
                if 'Direction:' in line:
                    direction = find_direction(line)
            elif not headings_found:
                if 'Time  Total   Cls   Cls   Cls   Cls   Cls   Cls   Cls   Cls   Cls   Cls   Cls   Cls' in line:
                    headings_found = True
            else:
                line_str = replace_multiple_spaces(line)
                first_token = str(line_str.split(' ')[0])
                if len(first_token) == 4:
                    if first_token == start_time:
                        analysis_date = analysis_date + timedelta(days=1)
                    new_row = [site, direction, analysis_date.strftime('%Y-%m-%d')]
                    new_row = new_row + line_str.split(' ')
                    df.loc[len(df)] = new_row
    return df


def find_all_files(folder, file_type=None):
    main_df = pd.DataFrame()
    if file_type is None:
        file_type = '.txt'
    #if folder
    all_files = os.listdir(folder)
    for file in tqdm(all_files, desc='analysing tube count files'):
        if file.endswith(file_type):
            file_path = f"{folder}\{file}"
            df_file = obtain_tube_data(file_path)
            main_df = pd.concat([main_df, df_file])
    return main_df


#df_site = obtain_tube_data(input_file)
df_tubes = find_all_files(input_folder, file_type='.txt')
df_tubes.to_csv(output_file, index=False)
print(df_tubes)
