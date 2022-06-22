import pandas as pd
import numpy as np
from tqdm import tqdm
from datetime import datetime, timedelta
import os
from gis import osm_tools
from IPython.display import display
from project_tools.intersection_turn_counts import file_utls as fu

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
    start_time = int(start_time.replace(':', ''))
    start_time = f"{start_time:04}"
    return start_time


def find_start_date(date_str):
    start_date = datetime.strptime(date_str, '%H:%M %A, %d %B %Y')
    return start_date


def obtain_tube_data(file_name, street_address):
    site = None
    direction = None
    date_strings = None
    headings_found = False
    headings = None
    df = None
    with open(file_name, "r") as f:
        for line in f:
            if site is None:
                if "Site:" in line:
                    site = find_site(line)
                    site = site.split('] ')[1]
                    street_address = site
                    # ToDo: use better tools to determine sfind site to search if not found previously
                    '''
                    address_split = line.replace(',', '').replace('no.', '').split(' ')
                    street_number = find_street_number(address_split)
                    suburb = find_suburb(address_split)
                    street_name = find_street_name(address_split, street_number, suburb)
                    street_address = f'{street_number} {street_name}, {suburb}, QLD, Australia'
                    '''
            elif date_strings is None:
                if "Filter time:" in line:
                    date_strings = find_date_range(line)
                    start_time = find_four_digit_start_time(date_strings[0])
                    analysis_date = try_parsing_date(date_strings[0])
            elif headings is None:
                if 'Included classes:' in line:
                    headings = determine_headings(line)
                    df = pd.DataFrame(columns=headings)
            elif direction is None:
                if 'Direction:' in line:
                    direction = find_direction(line)
            elif not headings_found:
                if 'Time  Total   Cls' in line:
                    headings_found = True

            else:
                line_str = replace_multiple_spaces(line)
                first_token = str(line_str.split(' ')[0])
                if len(first_token) == 4:
                    if first_token == start_time:
                        analysis_date = analysis_date + timedelta(days=1)
                    new_row = [site, direction, analysis_date.strftime('%Y-%m-%d')]
                    new_row = new_row + line_str.split(' ')
                    df.loc[len(df)] = new_row[:len(df.columns)]# quick fix to exclude speed data if it is tagged on at the end

    if df is None:
        new_row = [site, direction, date_strings, None, None, None, None, None, None, None, None, None, None, None,
                   None, None, None]
        headings = ['site', 'direction', 'date', 'time', 'total', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                    '11', '12']
        df = pd.DataFrame(columns=headings)
        df.loc[len(df)] = new_row[:len(df.columns)]
    df = df.melt(id_vars=['site', 'direction', 'date', 'time', 'total'], var_name='vehicle_class', value_name='count')
    return df, street_address


def find_suburb(address_split):
    suburb = None
    for i in address_split:
        if i == i.upper() and not i.isnumeric():
            if suburb is None:
                suburb = i.strip()
            else:
                suburb += f' {i.strip()}'
    return suburb


def find_street_number(address_split):
    street_number = None
    if address_split[0].isnumeric():
        street_number = address_split[0]
    else:
        numeric_strings = 0
        for i in address_split[0:]:
            i
            if i.isnumeric():
                street_number = i
                numeric_strings += 1
        if numeric_strings != 1:
            street_number = 'not identified'
    return street_number


def find_street_name(address_split, street_number, suburb):
    street_name = None
    suburb_found = False
    if suburb is not None:
        for i in address_split:
            if i == street_number:
                continue
            elif i == suburb:
                suburb_found = True
                break
            else:
                if street_name is None:
                    street_name = i
                else:
                    street_name += f' {i}'
    if not suburb_found:
        street_name = None
    return street_name


def address_from_file(file_name):
    address = file_name.replace("Outside no.", "")
    address = address.replace("#", "")
    address = address.split('_', 1)[1]
    address = address.split(' Class Volume', 1)[0]
    address_split = address.split(' ')
    street_number = find_street_number(address_split)
    suburb = find_suburb(address_split)
    street_name = find_street_name(address_split, street_number, suburb)
    street_address = f'{street_number} {street_name}, {suburb}, Queensland'
    return street_address


def address_from_file_v2(file_name):
    address = file_name.replace("Outside no.", "")
    address = address.replace("#", "")
    address = address.split('_', 1)
    if len(address) >=2:
        address = address[1]
        address = address.replace('_NB_', '____').replace('_EB_', '____').replace('_WB_', '____').replace('_SB_', '____')
        address_split = address.split('____')[0]
        street_address = address_split.replace('_', ' ')
    else:
        street_address = None
    return street_address


def analyse_files_in_folder(folder, output_file, file_type=None):
    #main_df = pd.DataFrame()
    if file_type is None:
        file_type = '.txt'
    all_files = os.listdir(folder)
    if fu.check_file_exists(output_file):
        all_files = fu.exclude_files_already_assessed(all_files=all_files, assessed_file=output_file)
    for file in tqdm(all_files, desc='analysing tube count files'):
        print(file)
        if file.endswith(file_type):
            file_path = f"{folder}\\{file}"
            street_address = address_from_file_v2(replace_multiple_spaces(file))
            df_file, street_address = obtain_tube_data(file_path, street_address)
            if street_address is not None or df_file.empty:
                lat, lon, location = osm_tools.geocode_coordinates(street_address, api='here')
                if lat is None:
                    lat, lon, location = osm_tools.geocode_coordinates(street_address, api='google')
            else:
                lat = None
                lon = None
                location = None
            # ToDo: add location to dataframe for comparison to street address.
            df_file.loc[:, 'lat'] = lat
            df_file.loc[:, 'lon'] = lon
            df_file.loc[:, 'street_address'] = street_address
            df_file.loc[:, 'file_name'] = file
            fu.save_dataframe_log(df_file, output_file)
            #main_df = pd.concat([main_df, df_file])
    return #main_df


def try_parsing_date(text):
    for fmt in ('%H:%M %A, %d %B %Y', '%H:%M %A, %d %B, %Y'):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError('no valid date format found')