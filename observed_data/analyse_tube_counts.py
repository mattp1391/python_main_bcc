import pandas as pd
import numpy as np
import geopandas as gpd
from tqdm import tqdm
from datetime import datetime, timedelta
import os
from IPython.display import display
from project_tools.intersection_turn_counts import file_utls as fu
from project_tools.intersection_turn_counts import dataframe_utls as dfu
from project_tools.intersection_turn_counts import gis_tools as gis


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
                    if street_address is None:
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
                    df.loc[len(df)] = new_row[
                                      :len(df.columns)]  # quick fix to exclude speed data if it is tagged on at the end
    if df is not None:
        df = df.melt(id_vars=['site', 'direction', 'date', 'time'], var_name='vehicle_class', value_name='count')
        df = pd.pivot_table(df, values='count', columns='time',
                            index=['site', 'direction', 'date', 'vehicle_class'])
        df = dfu.flatten_multi_index_columns_from_pivot(df, join_character='|', level=None)
    '''
    else:
        new_row = [site, direction, date_strings, None, None, None, None, None, None, None, None, None, None, None,
                   None, None, None]
        headings = ['site', 'direction', 'date', 'time', 'total', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                    '11', '12']
        df = pd.DataFrame(columns=headings)
        df.loc[len(df)] = new_row[:len(df.columns)]
    '''
    return df, street_address, direction


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
    if len(address) >= 2:
        address = address[1]
        address = address.replace('_NB_', '____').replace('_EB_', '____').replace('_WB_', '____').replace('_SB_',
                                                                                                          '____')
        address_split = address.split('____')[0]
        street_address = address_split.replace('_', ' ')
    else:
        street_address = None
    return street_address


def analyse_files_in_folder(folder, output_file, file_type=None):
    # main_df = pd.DataFrame()
    offset_left_dict = {'north': 270,
                        'east': 0,
                        'south': 90,
                        'west': 180,
                        }
    if file_type is None:
        file_type = '.txt'
    all_files = os.listdir(folder)
    if fu.check_file_exists(output_file):
        all_files = fu.exclude_files_already_assessed(all_files=all_files, assessed_file=output_file)
    for file in tqdm(all_files, desc='analysing tube count files'):
        # print(file)
        lat = None
        lon = None
        location = None
        if file.endswith(file_type):
            file_path = f"{folder}\\{file}"
            street_address = address_from_file_v2(replace_multiple_spaces(file))
            tube_info = obtain_tube_data(file_path, street_address)
            df_file = tube_info[0]
            street_address = tube_info[1]
            direction = tube_info[2]
            if df_file is not None:
                if street_address is not None:
                    geo_search_text = gis.add_qld_aus_to_geolocate_text(street_address)
                    lat, lon, location = gis.geocode_coordinates(geo_search_text, api='here')
                    if lat is None:
                        lat, lon, location = gis.geocode_coordinates(geo_search_text, api='google')
                # ToDo: add location to dataframe for comparison to street address.
                direction_angle = offset_left_dict.get(direction.lower())
                if direction_angle is not None:
                    offset_point = gis.find_point_at_distance_and_bearing(lat, lon, distance=0.005,
                                                                          bearing=direction_angle)
                    lat = offset_point[0]
                    lon = offset_point[1]
                df_file.loc[:, 'date'] = pd.to_datetime(df_file['date'])
                # movement_log_df.loc[:, 'survey_date'] = pd.to_datetime(movement_log_df['survey_date'])
                df_file.loc[:, 'day'] = np.where(df_file['date'].isnull(), np.nan, df_file['date'].dt.day_name())
                # df_file.loc[:, 'day'] = np.where(df_file['date'].isnull(), np.nan, df_file['date'].dt.day_name())
                df_file.loc[:, 'lat'] = lat
                df_file.loc[:, 'lon'] = lon
                df_file.loc[:, 'street_address'] = geo_search_text
                df_file.loc[:, 'file_name'] = file
                expected_headers = [
                    "site", "direction", "date", "day", "lat", "lon", "street_address", "vehicle_class", "0000", "0015",
                    "0030", "0045", "0100", "0115", "0130", "0145", "0200", "0215", "0230", "0245", "0300", "0315",
                    "0330", "0345", "0400", "0415", "0430", "0445", "0500", "0515", "0530", "0545", "0600", "0615",
                    "0630", "0645", "0700", "0715", "0730", "0745", "0800", "0815", "0830", "0845", "0900", "0915",
                    "0930", "0945", "1000", "1015", "1030", "1045", "1100", "1115", "1130", "1145", "1200", "1215",
                    "1230", "1245", "1300", "1315", "1330", "1345", "1400", "1415", "1430", "1445", "1500", "1515",
                    "1530", "1545", "1600", "1615", "1630", "1645", "1700", "1715", "1730", "1745", "1800", "1815",
                    "1830", "1845", "1900", "1915", "1930", "1945", "2000", "2015", "2030", "2045", "2100", "2115",
                    "2130", "2145", "2200", "2215", "2230", "2245", "2300", "2315", "2330", "2345", "file_name"]
                df_columns_sorted = sorted(df_file.columns.tolist())
                expected_headers_sorted = sorted(expected_headers)
                if expected_headers_sorted == df_columns_sorted:
                    df_file = df_file[expected_headers]
                    fu.save_dataframe_log(df_file, output_file)
                else:
                    print(f"Headers do not match for {file}.  Please check data input.")
            else:
                print(f'no data found for {file}.')
    return


def try_parsing_date(text):
    for fmt in ('%H:%M %A, %d %B %Y', '%H:%M %A, %d %B, %Y'):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError('no valid date format found')


def create_geojson_from_tube_analysis_csv(csv_file, geo_json_file):
    df = pd.read_csv(csv_file, encoding='cp1252')
    weekend_days = ['saturday', 'sunday', 'sat', 'sun']
    df_filtered = df[(~df['day'].str.lower().isin(weekend_days)) & (df['vehicle_class'].str.lower() == 'total')]
    list_of_sum_columns = ["0000", "0015", "0030", "0045", "0100", "0115", "0130", "0145", "0200", "0215", "0230",
                           "0245", "0300", "0315", "0330", "0345", "0400", "0415", "0430", "0445", "0500", "0515",
                           "0530", "0545", "0600", "0615", "0630", "0645", "0700", "0715", "0730", "0745", "0800",
                           "0815", "0830", "0845", "0900", "0915", "0930", "0945", "1000", "1015", "1030", "1045",
                           "1100", "1115", "1130", "1145", "1200", "1215", "1230", "1245", "1300", "1315", "1330",
                           "1345", "1400", "1415", "1430", "1445", "1500", "1515", "1530", "1545", "1600", "1615",
                           "1630", "1645", "1700", "1715", "1730", "1745", "1800", "1815", "1830", "1845", "1900",
                           "1915", "1930", "1945", "2000", "2015", "2030", "2045", "2100", "2115", "2130", "2145",
                           "2200", "2215", "2230", "2245", "2300", "2315", "2330", "2345"]
    df_filtered.loc[:, 'daily'] = df_filtered[list_of_sum_columns].sum(axis=1)
    df_filtered = df_filtered[
        ['site', 'direction', 'lat', 'lon', 'street_address', 'vehicle_class', 'file_name', 'daily']]
    # display(df_filtered.info(verbose=True, null_counts=True))
    sites = df_filtered['site'].unique().tolist()

    geo_frames = []
    for site in tqdm(sites):
        mini_grouped_df = dfu.groupby_for_filtered_frame(df_filtered, 'site', site,
                                                         group_by_cols=['site', 'direction', 'lat', 'lon',
                                                                        'street_address', 'vehicle_class', 'file_name'])
        geo_frames.append(mini_grouped_df)
    gdf = pd.concat(geo_frames)
    gdf = gpd.GeoDataFrame(gdf, geometry=gpd.GeoSeries.from_xy(gdf['lon'], gdf['lat']), crs=4326)
    gdf.to_file(geo_json_file, driver="GeoJSON")
    return

