import math
import sys
import time

import contextily as cx
import geopandas as gpd
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from PIL import Image
from geopy import geocoders, Point
from geopy.distance import distance as geopy_distance
from matplotlib.backends.backend_agg import FigureCanvasAgg
from IPython.display import display
from shapely import wkt
from shapely.geometry import LineString
from datetime import datetime as dt

script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path: sys.path.append(script_folder)
from project_tools.intersection_turn_counts import file_utls
from importlib import reload

reload(file_utls)

google_geo_code_key = 'AIzaSyDgFypRMb-gnE9eaFjiWjcdc6T4JpjGUAo'
proxies = {'http': 'http://165.225.226.22:10170',
           'https': 'http://165.225.226.22:10170'}


def geocode_coordinates(address, user_agent='Engineering_Services_BCC', api='osm'):
    lat_ = None
    long_ = None
    location_ = None

    if api.lower() == 'osm':
        app = geocoders.Photon(user_agent=user_agent, proxies='165.225.226.22:10170')
        time.sleep(2)
    elif api.lower() == 'here':
        app = geocoders.HereV7(user_agent=user_agent, proxies='165.225.226.22:10170',
                               apikey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ')
        time.sleep(2)
    elif api.lower() == 'arcgis':
        app = geocoders.ArcGIS(user_agent=user_agent, proxies='165.225.226.22:10170')
        time.sleep(2)
    elif api.lower() == 'google':
        app = geocoders.GoogleV3(user_agent=user_agent, proxies='165.225.226.22:10170', api_key=google_geo_code_key)
        time.sleep(2)
    # print(address, api)
    location_ = app.geocode(address)

    if location_ is not None:
        lat_ = location_.latitude
        long_ = location_.longitude
    else:
        lat_ = None
        long_ = None
    return lat_, long_, location_


'''
location = 'Ann Street Between Roma St and Edward St BRISBANE_CITY'
lat, long, location_out = geocode_coordinates(location, api='osm')
print('osm', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='here')
print('here', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='arcgis')
print('arcgis', lat, long, location_out)

lat, long, location_out = geocode_coordinates(location, api='google')
print('google', lat, long, location_out)
'''


def define_bounding_box(lat_min, lat_max, lon_min, lon_max):
    b_box = (lon_min, lon_max, lat_min, lat_max)
    return b_box


def find_point_at_distance_and_bearing(lat, lon, distance=1.0, bearing=0):
    """
    Calculate destination point from a starting point (lat and lon), distance and bearing.

    Parameters
    ----------
    lat (float): Latitude of starting point
    lon (float): longitude of starting point
    distance (int): Distance in kilomters
    bearing (float): bearing in degrees to destination point.  N = 0, E = 90, S = 180, W = 270 or -90.

    Returns
    -------
    Point: Point(lat, lon, altitude)
    """
    origin = Point(lat, lon)
    destination = geopy_distance(kilometers=distance).destination(origin, bearing)
    return destination


def angle_between(p1, p2):
    """
    find angle of a point.  NOTE THIS HAS BEEN SUPERCEDED!
    Parameters
    ----------
    p1
    p2

    Returns
    -------

    """
    ang1 = np.arctan2(*p1[::-1])
    ang2 = np.arctan2(*p2[::-1])
    radians = (ang2 - ang1) % (2 * np.pi)
    degrees = np.rad2deg(radians)
    return degrees


def compass_angle(p1, p2, excel_cell_format=False):  # updated to iinclude as x and y for points individually
    """
    find angle of a point.  NOTE THIS HAS BEEN DEPRECATED!
    Parameters
    ----------
    p1(array or list): point with (x, y) or [x, y]
    p2(array or list): point with (x, y) or [x, y]

    Returns
    -------

    """
    # ToDo: update docstring

    if excel_cell_format:
        origin_x = p1[1]
        origin_y = p1[0]
        destination_x = p2[1]
        destination_y = p2[0]
    else:
        origin_x = p1[0]
        origin_y = p1[1]
        destination_x = p2[0]
        destination_y = p2[1]
    delta_x = destination_x - origin_x
    delta_y = destination_y - origin_y
    degrees_temp = math.atan2(delta_x, delta_y) / math.pi * 180
    if degrees_temp < 0:
        degrees_final = degrees_temp + 360
    else:
        degrees_final = degrees_temp
    return degrees_final


def compass_angle_ss(origin_x, origin_y, destination_x, destination_y):
    """
    find the angle of two points relative to north.
    Parameters
    ----------
    origin_x (float): number representing latitude of first point
    origin_y (float): number representing longitude of first point
    destination_x (float): number representing latitude of second point
    destination_y (float): number representing longitude of second point

    Returns
    -------
    float: number from 0 - 360 representing angle from north.
    """
    delta_x = destination_x - origin_x
    delta_y = destination_y - origin_y
    degrees_temp = math.atan2(delta_x, delta_y) / math.pi * 180
    if degrees_temp < 0:
        degrees_final = degrees_temp + 360
    else:
        degrees_final = degrees_temp
    return degrees_final


def custom_round(x, base=5):
    """
    round any number to the base number
    Parameters
    ----------
    x (float): number to be rounded
    base (float): number to be rounded to

    Returns
    -------

    """
    return base * round(x / base)


def closest_node(node, nodes):
    """
    find closest point in a list of points (nodes)
    Parameters
    ----------
    node
    nodes

    Returns
    -------

    """
    nodes = nodes.remove(node)
    nodes = np.asarray(nodes)
    dist_2 = np.sum((nodes - node) ** 2, axis=1)
    return np.argmin(dist_2)


def wkt_loads(wkt_col):
    """
    converts wkt into geopandas geometry.
    Parameters
    ----------
    wkt_col

    Returns
    -------
    geometry column from wkt col
    """
    try:
        return wkt.loads(wkt_col)
    except Exception:
        return None


def find_point_in_linestring(line_str, point_index=0, gdf_point=True):
    coordinates = line_str.coords
    # if len(coordinates)==1:
    #    print (1)
    point = coordinates[point_index]
    # lat = point[0]
    # lon = point[1]
    # wkt_point = f'POINT ({lat} {lon})'

    return point


def find_angle_of_linestring(linestring, point_1_index=0, point_2_index=-1):
    line_str_coordinates = linestring.coords
    point_1 = line_str_coordinates[point_1_index]
    point_2 = line_str_coordinates[point_2_index]
    angle = compass_angle(point_1, point_2)
    return angle


def add_direction_columns_to_gdf(gdf, geometry_col='geometry', angle_col='angle'):
    # ToDo: check if new columns are in dataframe before adding them.
    gdf.loc[:, angle_col] = gdf.apply(lambda row: find_angle_of_linestring(row[geometry_col]), axis=1)
    gdf.loc[:, 'angle_round_90'] = gdf.apply(lambda row: custom_round(row[angle_col], 90), axis=1)
    gdf.loc[:, 'angle_round_45'] = gdf.apply(lambda row: custom_round(row[angle_col], 45), axis=1)
    direction_dict = get_direction_dict()
    gdf.loc[:, 'direction_4'] = gdf['angle_round_90'].map(direction_dict)
    gdf.loc[:, 'direction_8'] = gdf['angle_round_45'].map(direction_dict)
    return gdf


def get_direction_dict():
    dict_ = {0: 'N',
             45: 'NE',
             90: 'E',
             135: 'SE',
             180: 'S',
             225: 'SW',
             270: 'W',
             315: 'NW',
             360: 'N',
             -45: 'NW',
             -90: 'W',
             -135: 'SW',
             -180: 'S'
             }
    return dict_


def plot_gdf_with_map(gdf_link, gdf_node, label='NodeId', colour_col=None, cmap=None, marker_size=300,
                      text_size='text_size'):
    gdf_with_map = gdf_link.to_crs(epsg=3857)
    gdf_node_map = gdf_node.to_crs(epsg=3857)
    ax = gdf_with_map.plot(colour_col, cmap=cmap, figsize=(10, 10), alpha=0.5, linewidth=10.0)
    gdf_node_map.plot(ax=ax, markersize=marker_size, c='black')

    ax.plot()
    for x, y, label, text_size in zip(gdf_node_map.geometry.x, gdf_node_map.geometry.y, gdf_node_map[label],
                                      gdf_node_map[text_size]):
        ax.annotate(label, xy=(x, y), xytext=(3, 3), textcoords="offset points", fontsize=text_size)
    ax.set_axis_off()
    # for x, y, label in zip(gdf_link['geometry'].x, gdf_link['geometry'].y, gdf_link['label']):
    #    ax.annotate(label, xy=(x, y), xytext=(3, 3), textcoords="offset points")
    # cx.add_basemap(ax, source=cx.providers.Esri.WorldImagery)
    cx.add_basemap(ax, source=cx.providers.HEREv3.satelliteDay(apiKey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ',
                                                               proxies='165.225.226.22:10170'))
    cx.add_basemap(ax,
                   source=cx.providers.HEREv3.mapLabels(apiKey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ', size=128,
                                                        proxies='165.225.226.22:10170'))
    canvas = plt.get_current_fig_manager().canvas
    agg = canvas.switch_backends(FigureCanvasAgg)
    s, (width, height) = agg.print_to_buffer()
    im = Image.frombytes("RGBA", (width, height), s)
    return im


def create_sections_gdf(sections_file, crs=None):
    gdf = gpd.read_file(sections_file)
    gdf.columns = gdf.columns.str.lower()
    # isplay(gdf.head())
    gdf = gdf.astype({'anode': 'str', 'bnode': 'str', 'cnode': 'str'})
    return gdf


def create_nodes_gdf(nodes_file, crs='epsg:4326'):
    gdf = gpd.read_file(nodes_file, crs=crs)
    gdf = gdf.astype({'NodeId': 'str'})
    return gdf


def find_node_distance_from_intersection(nodes_gdf, lat=None, lon=None, survey_df=None):
    if lat is None and lon is None:
        if survey_df is not None:
            intersection_gdf = survey_df[['intersection', 'geometry']].drop_duplicates().to_crs(crs=nodes_gdf.crs)
    else:
        # ToDo: Create geodataframe from lat and lon
        print('NEED TO CREATE GEOPDATAFRAME FROM LAT< LON PROVIDED')
    joined_gdf = gpd.sjoin_nearest(nodes_gdf, intersection_gdf, how='inner', distance_col='join_distance')
    return joined_gdf


def get_intersection_node(df):
    # display(df.head(), 2)
    # display(df[df['join_distance'] == df['join_distance'].min()]['NodeId'])
    if df.empty:
        closest_node_from_df = None
        distance_from_node = None
    else:
        closest_node_from_df = df[df['join_distance'] == df['join_distance'].min()]['NodeId'].iloc[0]
        distance_from_node = df['join_distance'].min()
    return closest_node_from_df, distance_from_node


def find_intersection_links(sections_gdf, intersection_links, nodes):
    links_from_gdf = sections_gdf[sections_gdf['b_node'].isin(nodes)]
    links_from_gdf.loc[:, 'approach_type'] = 'from'
    links_to_gdf = sections_gdf[sections_gdf['a_node'].isin(nodes)]
    links_to_gdf.loc[:, 'approach_type'] = 'to'
    links_gdf = pd.concat([links_from_gdf, links_to_gdf])
    '''
    links_gdf.loc[:, 'angle'] = links_gdf.apply(
        lambda row: compass_angle((row['AXco'], row['AYco']), (row['BXco'], row['BYco']), False), axis=1)
    links_gdf.loc[:, 'angle_round_8'] = links_gdf.apply(lambda row: custom_round(row['angle'], 45),
                                                                        axis=1)
    links_gdf.loc[:, 'direction_8'] = links_gdf['angle_round'].map(get_direction_dict)
    links_gdf.loc[:, 'angle_round_4'] = links_gdf.apply(lambda row: custom_round(row['angle'], 90),
                                                                        axis=1)
    links_gdf.loc[:, 'direction_4'] = links_gdf['angle_round'].map(get_direction_dict)
    '''
    links_gdf.loc[:, 'colour'] = np.where(links_gdf['link_a_b'].isin(intersection_links), 0, 1)
    return links_gdf


def sort_dictionary(dictionary):
    keys = dictionary.keys()
    keys.sort(key=lambda k: (k[0], int(k[1:])))
    return map(dictionary.get, keys)


def add_to_log(xl_file, log_type, df=None, movements=None, comments=None):
    if log_type == 1:
        log_file = r"D:\MP\projects\bcasm\log files\files_analysed.txt"
        df['ijk'] = df['spreadsheet_movement'].map(movements)
        df[['i', 'j', 'k']] = df['ijk'].str.split("_", expand=True)
        file_utls.create_csv_output_file(df, xl_file, output_folder=None)
        with open(log_file, "a") as file_object:
            file_object.write(f"{xl_file}, {comments}")
    else:
        log_file = r"D:\MP\projects\bcasm\log files\files_not_analysed.txt"
        with open(log_file, "a") as file_object:
            file_object.write(f"\n{xl_file}, {log_type}, {comments}")
    return


def create_map_image(add_to_database, ijk_movements, xl_file,
                     survey_df):  # int_node, excel_file_path, sections_gdf, nodes_gdf, dist_within=150):
    if add_to_database == 0:
        # ToDo add this to the log of unknown counts
        add_to_log(xl_file, add_to_database, comments="approaches don't match")
    elif add_to_database == 1:
        # ToDo add this to the database and log of known counts
        print('add to database and log')
        add_to_log(xl_file, add_to_database, comments=None)

    elif add_to_database == 2:
        log_for_later = True
        if log_for_later:
            add_to_log(xl_file, add_to_database, comments='use_map_for_this')

        else:
            # ToDo: add coded below
            print('fix this later')
            '''
            add_to_log(log_file, xl_file, log_type, comments=None)
            links_from_df = sections_gdf[(sections_gdf['BNode'] == int_node)]
            links_to_df = sections_gdf[(sections_gdf['ANode'] == int_node)]
            links_df = pd.concat([links_from_df, links_to_df])
            intersection_links = links_df['link_a_b'].unique().tolist()
            nodes = nodes_gdf[nodes_gdf['join_distance'] <= dist_within]['NodeId'].unique().tolist()
            links_gdf = find_intersection_links(sections_gdf, intersection_links, nodes)

            nodes_to_map = links_gdf['ANode'].unique().tolist()
            nodes_to_map.append(links_gdf['BNode'].unique().tolist())
            nodes_to_map = nodes_gdf[nodes_gdf['NodeId'].isin(nodes_to_map)]
            nodes_to_map.loc[:, 'marker_size'] = np.where(nodes_to_map['NodeId'] == int_node, 300, 150)
            nodes_to_map.loc[:, 'text_size'] = np.where(nodes_to_map['NodeId'] == int_node, 20, 10)

            cmap = LinearSegmentedColormap.from_list('mycmap', [(0, 'red'), (1, 'grey')])
            #plot_gdf_with_map(gdf_link, gdf_node, label='NodeId', colour_col=None, cmap=None, marker_size=300,
            #                  text_size='text_size'):

            map_image = plot_gdf_with_map(links_gdf, nodes_to_map, cmap=cmap, label='NodeId', colour_col='colour',
                                          marker_size='marker_size', text_size='text_size')
            xl_img = aic.get_excel_image(file_path=excel_file_path, sheet_name='Sheet1', range=None)
            '''
    return  # map_image


def filter_direction(gdf, direction_col, direction):
    movement_gdf = gdf[(gdf[direction_col] == direction)]
    return movement_gdf


def find_geometry_of_link_node(df, node_col, nodes_gdf):
    df = df[[node_col]]
    df['lon'] = df.apply(lambda x: x[node_col][0], axis=1)
    df['lat'] = df.apply(lambda x: x[node_col][1], axis=1)
    df = df.drop_duplicates()
    gdf = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df['lon'], df['lat']))
    joined_df = gpd.sjoin_nearest(gdf, nodes_gdf, how='inner', distance_col='join_distance_a')
    joined_df = joined_df[[node_col, 'NodeId']]
    return joined_df


def reverse_coordinates(line_str):
    coordinates = line_str.coords
    coordinates = coordinates[::-1]
    return coordinates


def create_linestring(coordinates):
    if len(coordinates) > 1:
        line_str = LineString(coordinates)

    else:
        line_str = 'delete me'
    return line_str


def find_node_start_and_end(network_links_gdf, nodes_gdf, start_node_col='a_node', end_node_col='b_node',
                            link_id='link_a_b'):
    """

    Parameters
    ----------
    sections_gdf ( geodataframe): geodataframe containing network links
    nodes_gdf (geodataframe): geodataframe containing node points

    Returns
    -------
    sections_gdf geodataframe with 3 additional columns which contain the start node, end node and link_id
    ('start_node_end_node').
    """
    network_links_gdf.loc[:, start_node_col] = network_links_gdf.apply(
        lambda row: find_point_in_linestring(row['geometry'], point_index=0), axis=1)

    network_links_gdf.loc[:, end_node_col] = network_links_gdf.apply(
        lambda row: find_point_in_linestring(row['geometry'], point_index=-1), axis=1)
    a_joined = find_geometry_of_link_node(network_links_gdf, start_node_col, nodes_gdf)
    output_gdf = pd.merge(network_links_gdf, a_joined[[start_node_col, 'NodeId']])
    output_gdf = output_gdf.rename(columns={'NodeId': 'a_node_id'})
    b_joined = find_geometry_of_link_node(network_links_gdf, end_node_col, nodes_gdf)
    output_gdf = pd.merge(output_gdf, b_joined[[end_node_col, 'NodeId']])
    output_gdf = output_gdf.rename(columns={'NodeId': 'b_node_id'})
    output_gdf = output_gdf[['geometry', 'a_node', 'b_node', 'a_node_id', 'b_node_id']]
    # output_gdf_2 = output_gdf.copy()
    output_gdf_2 = output_gdf.rename(columns={'a_node': 'b_node', 'b_node': 'a_node', 'a_node_id': 'b_node_id',
                                              'b_node_id': 'a_node_id'})
    # output_gdf_2['coordinates_reversed'] = output_gdf_2['geometry'].coords.reverse()
    # arr.reverse()

    # tqdm.pandas()
    output_gdf_2['coordinates_reversed'] = output_gdf_2.apply(lambda row: reverse_coordinates(row['geometry']), axis=1)

    output_gdf_2['line_str_reversed'] = output_gdf_2.apply(
        lambda row: create_linestring(row['coordinates_reversed']), axis=1)  # Create a linestring column
    output_gdf_2 = output_gdf_2[output_gdf_2['line_str_reversed'] != 'delete me']
    output_gdf_2 = output_gdf_2.drop('geometry', axis=1)  # Drop WKT column
    output_gdf_2 = output_gdf_2.set_geometry(col='line_str_reversed')
    output_gdf_2 = output_gdf_2.rename(columns={'line_str_reversed': 'geometry'})
    output_gdf_2 = output_gdf_2[['geometry', 'a_node', 'b_node', 'a_node_id', 'b_node_id']]
    output_gdf = pd.concat([output_gdf, output_gdf_2])
    output_gdf.loc[:, link_id] = output_gdf['a_node_id'] + "_" + output_gdf['b_node_id']
    output_gdf = output_gdf.drop_duplicates(subset=link_id, keep="first")
    output_gdf = add_direction_columns_to_gdf(output_gdf)
    return output_gdf


def find_ijk(sections_gdf, nodes_gdf, survey_gdf, movement_dict=None):
    int_node, dist_to_node = get_intersection_node(nodes_gdf)
    dist_to_node = round(dist_to_node, 2)
    from_gdf = sections_gdf[(sections_gdf['b_node_id'] == int_node)]
    to_gdf = sections_gdf[(sections_gdf['a_node_id'] == int_node)]
    approach_to_direction = {'N': 'S', 'NE': 'SW', 'E': 'W', 'SE': 'NW', 'S': 'N', 'SW': 'NE', 'W': 'E', 'NW': 'SE'}
    i = []
    j = []
    k = []
    # movement_ijk_dict = {}
    # display('movement_ijk_dict', movement_ijk_dict)
    dist_to_node_list = []
    add_to_database = []
    direction_movements = []
    excel_keys = []
    angles_from = []
    angles_to = []
    for excel_key, movement in movement_dict.items():
        if int_node is not None:
            approaches = movement.split('_')
            from_8_approach_df = filter_direction(from_gdf, 'direction_8', approach_to_direction[approaches[0]])
            to_8_approach_df = filter_direction(to_gdf, 'direction_8', approaches[1])
            if len(from_8_approach_df) == 1 and len(to_8_approach_df) == 1:

                i.append(str(from_8_approach_df['a_node_id'].iloc[0]))
                j.append(str(int_node))
                k.append(str(to_8_approach_df['b_node_id'].iloc[0]))
                add_to_database.append(1)
                from_approach = approach_to_direction[from_8_approach_df['direction_8'].iloc[0]]
                to_approach = to_8_approach_df['direction_8'].iloc[0]
                from_angle = float(from_8_approach_df['angle'].iloc[0])
                to_angle = float(to_8_approach_df['angle'].iloc[0])
                angles_from.append(round(from_angle, 2))
                angles_to.append(round(to_angle, 2))
                direction_movements.append(f"{from_approach}_{to_approach}")
                # movement_ijk_dict[excel_key] = f'{i}_{j}_{k}'
            else:
                from_4_approach_df = filter_direction(from_gdf, 'direction_4', approach_to_direction[approaches[0]])
                to_4_approach_df = filter_direction(to_gdf, 'direction_4', approaches[1])
                if len(from_4_approach_df) == 1 and len(to_4_approach_df) == 1:
                    i.append(str(from_4_approach_df['a_node_id'].iloc[0]))
                    j.append(str(int_node))
                    k.append(str(to_4_approach_df['b_node_id'].iloc[0]))
                    from_approach = approach_to_direction[from_4_approach_df['direction_8'].iloc[0]]
                    to_approach = to_4_approach_df['direction_8'].iloc[0]
                    direction_movements.append(f"{from_approach}_{to_approach}")
                    from_angle = float(from_4_approach_df['angle'].iloc[0])
                    angle_90 = float(from_4_approach_df['angle_round_90'].iloc[0])
                    angle_from_dif = min(abs(angle_90 - from_angle), abs(angle_90 - from_angle + 360),
                                         abs(angle_90 - from_angle - 360))
                    # print(f"{from_approach}_{to_approach}")
                    # print('from_angle', from_angle, angle_90, angle_from_dif)
                    to_angle = float(to_4_approach_df['angle'].iloc[0])
                    angle_to_90 = float(to_4_approach_df['angle_round_90'].iloc[0])
                    angle_to_dif = min(abs(angle_to_90 - to_angle), abs(angle_to_90 - to_angle + 360),
                                       abs(angle_to_90 - to_angle - 360))
                    # print('to_angle', to_angle, angle_to_90, angle_to_dif)
                    angles_from.append(round(from_angle, 2))
                    angles_to.append(round(to_angle, 2))
                    if angle_to_dif <= 35 and angle_from_dif <= 35:
                        add_to_database.append(1.1)  # all approaches within 70 degrees of assumed direction.'
                    else:
                        add_to_database.append(1.2)
                else:
                    add_to_database.append(0.2)  # not a 1 is to 1 relationship found.
                    i.append(None)
                    j.append(None)
                    k.append(None)
                    angles_from.append(None)
                    angles_to.append(None)
                    direction_movements.append(None)
        else:
            add_to_database.append(0.3)  # not a 1 is to 1 relationship found.
            i.append(None)
            j.append(None)
            k.append(None)
            angles_from.append(None)
            angles_to.append(None)
            direction_movements.append(None)
        dist_to_node_list.append(dist_to_node)
        excel_keys.append(excel_key)
    movement_ijk_dict = {'excel_movement': excel_keys, 'geographic_movement': direction_movements,
                         'angle_from': angles_from, 'angle_to': angles_to, 'i': i, 'j': j, 'k': k,
                         'log_type': add_to_database, 'dist_to_node': dist_to_node}
    return movement_ijk_dict
