import pandas as pd
import numpy as np
import geopandas as gpd
from tqdm import tqdm
from datetime import date, datetime
import sys
from IPython.core.display import display

'''
Notes: 

Travel Direction - F = From Reference Node, T = To Reference Node, B = Both Directions


'''


def date_as_string(date_=None, date_format='%Y_%m_%d'):
    if date_ is None:
        date_ = date.today()
        date_str = date_.strftime(date_format)
    return date_str


def create_dataframe_from_geo_dataframe_if_geometry_empty(gdf, geometry_col='geometry'):
    if geometry_col in gdf.columns:
        geometry_values = gdf[geometry_col].values.tolist()
        if geometry_values.count(None) == len(geometry_values):
            df = pd.DataFrame(gdf.drop(columns='geometry'))
            return df
    return gdf


def import_gpkg_layers(file_name, layer_list):
    """
    creates a dictionary of dataframes from a geopeackage (".gpkg") file

    Parameters
    ----------
    file_name (string): location and name of .gpkg file package (eg. r"folder/filename.gpkg")
    layer_list (list): list of layers to be included in the dataframe dictionary

    Returns
    -------
    dictionary of geo-dataframes or dataframes consisting of the layer name and the relevant geo dataframe (if spatial) or pandas
    dataframe.
    """
    output_dict = {}
    for layer in tqdm(layer_list, desc=' loading layers from .gpkg file'):
        gdf = gpd.read_file(file_name, driver="GPKG", layer=layer)
        gdf = create_dataframe_from_geo_dataframe_if_geometry_empty(gdf)
        output_dict[layer] = gdf
    if output_dict == {}:
        output_dict = None
        print("no tables have been found or loaded to the dictionary")
    return output_dict

    # travel_dir_cond_rules = {1: ((gdf['TRAVEL_DIR'] == 'B') | (gdf['TRAVEL_DIR'] == 'F'))}
    # gdf['from_ref'] = np.where(((gdf['TRAVEL_DIR'] == 'B') | (gdf['TRAVEL_DIR'] == 'F')), 1, 0)
    # gdf['to_ref'] = np.where(((gdf['TRAVEL_DIR'] == 'B') | (gdf['TRAVEL_DIR'] == 'T')), 1, 0)
    # gdf['from_ref'] = np.select(travel_dir_cond_rules.values(), travel_dir_cond_rules.keys(), default=0)


def update_df_to_include_column_names_in_list(df, column_list):
    updated_cols = []
    for col in df.columns.tolist():
        if col in column_list:
            updated_cols.append(col)
    if not updated_cols:
        print('No columns found in filter list.  No cleaning undertaken for here data!')
    else:
        df = df[updated_cols]
    return df


def clean_here_2001_link(df):
    filter_cols = ['LINK_ID', 'LEFT_POSTAL_CODE', 'RIGHT_POSTAL_CODE', 'FUNCTIONAL_CLASS', 'TRAVEL_DIRECTION',
                   'SPEED_CATEGORY', 'FROM_REF_SPEED_LIMIT', 'TO_REF_SPEED_LIMIT', 'LENGTH', 'Shape_Length', 'T_F_DIR',
                   'JoinedVal', 'ROUTE_TYPE', 'ROAD_OWNER', 'Group Name', 'WARD', 'SUBURB_NAM', 'LINK_ID_TF',
                   'geometry']
    df = update_df_to_include_column_names_in_list(df, filter_cols)
    df = df.rename(columns={'JoinedVal': 'LINK_ID_TF'})
    df.loc[:, 'SPD_LIMIT_UPDT'] = np.where(df['T_F_DIR'] == "F", df['FROM_REF_SPEED_LIMIT'],
                                           df['TO_REF_SPEED_LIMIT'])
    return df


def join_pattern(df1, df2):
    df = df1.merge(df2, left_on='W', right_on='PATERN_ID')
    return df


def create_here_link_df(df_dict, road_types, road_type_col):
    df_here_link = clean_here_2001_link(df_dict['Here_2001_Link'])
    df_here_link = df_here_link[df_here_link[road_type_col].isin(road_types)]
    return df_here_link


def join_speed_data(df_dict, df_here_link):
    """
    This function finds the travel time ratio for all here links

    Parameters
    ----------
    df_dict (dictionary of dataframes): dictionary containing all dataframes required for assessment)
    road_type_col (str, optional): name of dataframe column with road types.  if None, 'ROUTE_TYPE' is used
    road_types (list, optional): list of road types to be assessed.  Default value will include all.

    Returns
    -------
    pd.DataFrame: Dataframe with speed data assessed.

    """
    df_ntp_ref_join = join_ntp_ref_oce_link(df_dict, df_here_link)
    df_here_speed = df_here_link.merge(df_ntp_ref_join, how='inner', on='LINK_ID_TF')
    df_joined = df_here_speed.merge(df_dict['NTP_SPD_OCE_60MIN_KPH_191H0'][['PATTERN_ID', 'H08_00', 'H17_00']],
                                    how='inner', left_on='W', right_on='PATTERN_ID')
    df_tt_ratios = find_tt_ratios(df_joined)
    return df_tt_ratios


def calc_mps(df, col_list):#  Don't think this is required.  Simpler way to calc ratio used.
    """

    Parameters
    ----------
    df
    col_list

    Returns
    -------

    """
    for col in col_list:
        new_col_name = f"{col}_MpS"
        df.loc[:, new_col_name] = df[col] * 1000 / 3600
    return df


def find_tt_ratios(df):
    # Todo: make functionality dynamic for different time periods.  Probably with dictionary.
    # ToDo: add parameters for column names
    """
    Find the travel time ratios for all here link data.  Ratio is calculated as speed limit divided by actual speed.
    At present functionality is hard coded to 8am for am and 5pm for PM.

    Parameters
    ----------
    df (pandas.DataFrame): dataframe consisting of link free flow speeds and actual speeds.

    Returns
    -------
    dataframe for export to csv
    """

    heading_filter = ['LINK_PVID', 'LINK_ID', 'TRAVEL_DIRECTION', 'LINK_ID_TF', 'H08_00', 'H17_00', 'LEFT_POSTAL_CODE',
                      'RIGHT_POSTAL_CODE', 'FUNCTIONAL_CLASS', 'TRAVEL_DIRECTION', 'SPEED_CATEGORY',
                      'FROM_REF_SPEED_LIMIT', 'TO_REF_SPEED_LIMIT', 'LENGTH', 'Shape_Length', 'T_F_DIR', 'ROUTE_TYPE',
                      'ROAD_OWNER', 'Group Name', 'WARD', 'SUBURB_NAM', 'SPD_LIMIT_UPDT', 'geometry']
    df = update_df_to_include_column_names_in_list(df, heading_filter)
    df = df.astype({'SPD_LIMIT_UPDT': 'float64', 'H08_00': 'float', 'H17_00': 'float'})
    df.loc[:, 'tt_ratio_am'] = df['SPD_LIMIT_UPDT'].div(df['H08_00'].values)
    df.loc[:, 'tt_ratio_pm'] = df['SPD_LIMIT_UPDT'].div(df['H17_00'].values)
    return df


def join_ntp_ref_oce_link(dataframes, ref_df):
    ref_df = ref_df.astype({'LINK_ID': 'float64'})
    if type(dataframes) is dict:
        dataframe_list = []
        for key, df in dataframes.items():
            if 'ntp_ref_oce_link' in key.lower():
                dataframe_list.append(df)
    elif type(dataframes) is list:
        dataframe_list = dataframes
    dataframes_to_concat = None
    for df in dataframe_list:
        df = update_df_to_include_column_names_in_list(df, column_list=['LINK_PVID', 'TRAVEL_DIRECTION', 'W'])
        df.loc[:, 'LINK_ID_TF'] = df['LINK_PVID'].astype(str) + df['TRAVEL_DIRECTION'].astype(str)
        df = df.astype({'LINK_PVID': 'float64'})
        df = filter_data_from_reference_dataframe(df=df, df_col='LINK_PVID', ref_df=ref_df, ref_df_col='LINK_ID')
        if dataframes_to_concat is None and df is not None:
            dataframes_to_concat = [df]
        elif df is not None:
            dataframes_to_concat.append(df)
    if dataframes_to_concat is not None:
        main_df = pd.concat(dataframes_to_concat)
    else:
        main_df = None
    return main_df


def filter_data_from_reference_dataframe(df, df_col, ref_df, ref_df_col):
    df_out = df.loc[df[df_col].isin(ref_df[ref_df_col].values.tolist()), :]
    if df_out.empty:
        df_out = None
    return df_out


def find_link_direction(lat_a, lon_a, lat_b, lon_b, t_f_dir):
    # f = from_reference
    # t = to reference
    ref_node = None
    if lat_a < lat_b:
        ref_node = None
    return ref_node


def create_intersection_centroids(df, owners_filter, owner_col, signal_filter, signal_col):
    df = df[df[owner_col].isin(owners_filter)]
    df = df[df[signal_col].isin(signal_filter)]
    df.loc[:, 'Buffer_m'] = df['Buffer_m']*2
    return df


def create_crash_data_df(df, date_column, start_date_str, end_date_str, date_format=None):
    if date_format is None:
        date_format = "%a %d-%b-%Y %I%p"
    df2 = df.copy()
    type_dict = {'Crash_Numb': str,
                 'WARD_CODE': str}
    df2 = df2.astype(type_dict)
    df2.loc[:, date_column] = df2[date_column].str.replace('am', 'AM')
    df2.loc[:, date_column] = df2[date_column].str.replace('pm', 'PM')
    #df.loc[:, date_column] = df2[date_column].dt.date
    start_date = datetime.strptime(start_date_str, '%d/%m/%Y').date()
    end_date = datetime.strptime(end_date_str, '%d/%m/%Y').date()
    df2.loc[:, date_column] = pd.to_datetime(df2[date_column], format=date_format).dt.date
    df2 = df2[(df2[date_column] > start_date) & (df2[date_column] <= end_date)]
    return df2


def get_here_10m(df_dict, here_10m_table=None):
    if here_10m_table is None:
        here_10m_name = 'Here_2001_Link_10m'

