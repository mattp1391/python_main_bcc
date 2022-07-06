import fiona
import math
import pandas as pd
import numpy as np
import geopandas as gpd
from tqdm import tqdm
import shapely
from shapely.geometry import LineString


def clip_geo_dataframe_from_large_file_on_import(file_path, boundary_gdf, chunk_size=50000, predicate_types=None,
                                                 layer=None):
    if predicate_types is None:
        #predicate_types = ['within', 'intersects']
        predicate_types = ['intersects']
    included_frames = []
    len_shape_file = len(fiona.open(file_path))
    chunks = list(range(0,len_shape_file, chunk_size))
    for i in tqdm(chunks):
        if layer is None:
            here_link = gpd.read_file(file_path, driver="fileGDB", rows=slice(i, min(i + chunk_size, len_shape_file)))
        else:
            here_link = gpd.read_file(file_path, driver="fileGDB", layer=layer,
                                  rows=slice(i, min(i+chunk_size, len_shape_file)))
        for predicate_type in predicate_types:
            df_joined = here_link.to_crs(epsg=4326).sjoin(boundary_gdf.to_crs(epsg=4326), how='inner',
                                                          predicate=predicate_type)
            if len(df_joined) > 0:
                included_frames.append(df_joined)
    print('Concatenating geo-dataframes.  This may take some time.')
    output_gdf = pd.concat(included_frames)
    output_gdf = output_gdf.drop_duplicates()
    return output_gdf


def find_lats_longs_dataframe(feature):
    lats = []
    lons = []
    linestrings = []
    if isinstance(feature, shapely.geometry.linestring.LineString):
        linestrings = [feature]
    elif isinstance(feature, shapely.geometry.multilinestring.MultiLineString):
        linestrings = feature.geoms
    for linestring in linestrings:
        x, y = linestring.xy
        lats = np.append(lats, y)
        lons = np.append(lons, x)
    return lats, lons


def find_ref_node(lats, lons):
    if lats[0] < lats[-1]:
        ref = 0
    elif lats[0] > lats[-1]:
        ref = 1
    elif lons[0] <= lons[-1]:
        # ToDo: technically if lats and lons are the same then check height z.
        ref = 0
    else:
        ref = 1
    return ref


def compass_angle(p1, p2, ref_node=0):  # updated to iinclude as x and y for points individually
    """
    find angle of a point. !
    Parameters
    ----------
    p1(array or list): point with (x, y) or [x, y]
    p2(array or list): point with (x, y) or [x, y]

    Returns
    -------

    """
    # ToDo: update docstring
    if ref_node == 0:
        origin_x = p1[0]
        origin_y = p1[1]
        destination_x = p2[0]
        destination_y = p2[1]
    else:
        origin_x = p1[1]
        origin_y = p1[0]
        destination_x = p2[1]
        destination_y = p2[0]
    delta_x = destination_x - origin_x
    delta_y = destination_y - origin_y
    degrees_temp = math.atan2(delta_x, delta_y) / math.pi * 180
    if degrees_temp < 0:
        degrees_final = degrees_temp + 360
    else:
        degrees_final = degrees_temp
    return degrees_final


def find_angle_of_linestring(linestring, link_id, point_1_index=0, point_2_index=-1):
    if len(linestring.geoms) > 1:
        print(link_id, linestring.geoms)
    linestring = linestring.geoms[0]
    line_str_coordinates = linestring.coords
    point_1 = line_str_coordinates[point_1_index]
    point_2 = line_str_coordinates[point_2_index]
    angle = compass_angle(point_1, point_2)
    return angle


def find_reverse_linestring_from_geometry(direction_type, geometry):
    line_str_reversed = None
    if direction_type.lower() in (['t', 'b']):
        line_str_reversed = LineString(list(geometry.geoms[0].coords)[::-1]).wkt
    return line_str_reversed


def create_link_in_opposite_direction_for_2_way_here_links(df):
    df_ref = df[df['TRAVEL_DIRECTION'].isin(['F', 'B'])]
    df_ref.loc[:, 'from_node'] = df_ref['REF_NODE_ID']
    df_ref.loc[:, 'to_node'] = df_ref['NON_REF_NODE_ID']
    df_ref.loc[:, 'f_t_direction'] = 'F'
    df_ref.loc[:, 'speed_limit_updated'] = df_ref['FROM_REF_SPEED_LIMIT']

    df_non_ref = df[df['TRAVEL_DIRECTION'].isin(['T', 'B'])]
    df_non_ref.loc[:,'line_str_reversed'] = df_non_ref.progress_apply(
        lambda row: find_reverse_linestring_from_geometry(row['TRAVEL_DIRECTION'], row['geometry']), axis=1)
    df_non_ref.loc[:, 'geometry'] = gpd.GeoSeries.from_wkt(df_non_ref['line_str_reversed'])
    df_non_ref.drop(columns='line_str_reversed')
    df_non_ref = pd.DataFrame(df_non_ref.drop(columns='geometry'))
    df_non_ref.loc[:, 'f_t_direction'] = 'T'
    df_non_ref.loc[:, 'from_node'] = df_non_ref['NON_REF_NODE_ID']
    df_non_ref.loc[:, 'to_node'] = df_non_ref['REF_NODE_ID']
    df_ref.loc[:, 'speed_limit_updated'] = df_ref['TO_REF_SPEED_LIMIT']

    gdf_final = pd.concat([df_ref, df_non_ref])
    gdf_final.loc[:, 'unique_id'] = gdf_final['LINK_ID'].astype(str) + gdf_final['f_t_direction'].astype(str)
    return gdf_final


def import_and_clean_ntp_ref_files(files, days=None):
    if isinstance(files, str):
        df = pd.read_csv(files)
    elif isinstance(files, list):
        df_list = []
        for f in files:
            df_list.append(pd.read_csv(f))
        df = pd.concat(df_list)
    if days is not None:
        if isinstance(days, list):
            cols = ['LINK_PVID', 'TRAVEL_DIRECTION'] + days
        elif isinstance(days, str):
            cols = ['LINK_PVID', 'TRAVEL_DIRECTION', days]
        df = df[cols]
    df.loc[:, 'unique_id'] = df['LINK_PVID'].astype(str) + df['TRAVEL_DIRECTION'].astype(str)
    return df


def import_and_clean_ntp_speed(file, time_filters=None):
    df = pd.read_csv(file)
    if time_filters is not None:
        if isinstance(time_filters, list):
            cols = ['PATTERN_ID'] + time_filters
        elif isinstance(time_filters, str):
            cols = ['PATTERN_ID', time_filters]
        df = df[cols]
    return df


def find_tt_ratios(df, speed_col='speed_limit_updated'):
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

    heading_filter = ['unique_id', 'LINK_ID', 'TRAVEL_DIRECTION', 'LINK_ID_TF', 'H08_00', 'H17_00', 'LEFT_POSTAL_CODE',
                      'RIGHT_POSTAL_CODE', 'FUNCTIONAL_CLASS', 'TRAVEL_DIRECTION', 'SPEED_CATEGORY',
                      'FROM_REF_SPEED_LIMIT', 'TO_REF_SPEED_LIMIT', 'LENGTH', 'Shape_Length', 'f_t_direction', 'ROUTE_TYPE',
                      'ROAD_OWNER', 'Group Name', 'WARD', 'SUBURB_NAM', speed_col, 'geometry']
    df = update_df_to_include_column_names_in_list(df, heading_filter)
    df = df.astype({speed_col: 'float64', 'H08_00': 'float', 'H17_00': 'float'})
    df.loc[:, 'tt_ratio_am'] = df[speed_col].div(df['H08_00'].values)
    df.loc[:, 'tt_ratio_pm'] = df[speed_col].div(df['H17_00'].values)
    return df


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
