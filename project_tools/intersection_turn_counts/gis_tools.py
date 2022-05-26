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

script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path: sys.path.append(script_folder)

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
    print(address, api)
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


def find_point_in_linestring(line_str, point_index=0):
    coordinates = line_str.coords
    point = coordinates[point_index]
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
    gdf = gdf.astype({'ANode': 'str', 'BNode': 'str', 'CNode': 'str'})
    gdf.loc[:, 'link_a_b'] = gdf['ANode'] + "_" + gdf['BNode']
    gdf = add_direction_columns_to_gdf(gdf)
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
    closest_node_from_df = df[df['join_distance'] == df['join_distance'].min()]['NodeId'].iloc[0]
    return closest_node_from_df


def find_intersection_links(sections_gdf, intersection_links, nodes):
    links_from_gdf = sections_gdf[sections_gdf['BNode'].isin(nodes)]
    links_from_gdf.loc[:, 'approach_type'] = 'from'
    links_to_gdf = sections_gdf[sections_gdf['ANode'].isin(nodes)]
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


def add_to_log(xl_file, log_type=0, log_file_assessed = True, comments=None):
    if log_file_assessed:
        log_file = r"D:\MP\projects\bcasm\log files\files_analysed.txt"
        with open(log_file, "a") as file_object:
            file_object.write(f"{xl_file}, {comments}")
    else:
        log_file = r"D:\MP\projects\bcasm\log files\files_not_analysed.txt"
        with open(log_file, "a") as file_object:
            file_object.write(f"{xl_file}, {log_type}, {comments}")


def create_csv_output_file(df, movements):
    print('create output code')



def create_map_image(add_to_database, ijk_movements, xl_file, survey_df):# int_node, excel_file_path, sections_gdf, nodes_gdf, dist_within=150):

    output_folder = r"D:\MP\projects\bcasm\log files\traffic_intersection_outputs"

    if add_to_database == 0:
        #ToDo add this to the log of unknown counts
        add_to_log(xl_file, add_to_database, comments="approaches don't match")
    elif add_to_database == 1:
        #ToDo add this to the database and log of known counts
        print('add to database and log')
        add_to_log(xl_file, add_to_database, comments=None)
        create_csv_output_file

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
    return #map_image


def filter_direction(gdf, direction_col, direction):
    movement_gdf = gdf[(gdf[direction_col] == direction)]
    return movement_gdf


def find_ijk(sections_gdf, nodes_gdf, movement_dict=None):
    int_node = get_intersection_node(nodes_gdf)
    from_gdf = sections_gdf[(sections_gdf['BNode'] == int_node)]
    to_gdf = sections_gdf[(sections_gdf['ANode'] == int_node)]
    add_to_database = 1 # value of zero will add to the log of unknowns.
    approach_to_direction = {'N': 'S', 'NE': 'SW', 'E': 'W', 'SE': 'NW', 'S': 'N', 'SW': 'NE', 'W': 'E', 'NW': 'SE'}
    i = None
    j = None
    k = None
    movement_ijk_dict = {}
    for excel_key, movement in movement_dict.items():

        approaches = movement.split('_')
        from_8_approach = filter_direction(from_gdf, 'direction_8', approach_to_direction[approaches[0]])
        to_8_approach = filter_direction(to_gdf, 'direction_8', approaches[1])
        if len(from_8_approach) == 1 and len(to_8_approach) == 1:
            i = from_8_approach['ANode'].iloc[0]
            j = int_node
            k = to_8_approach['BNode'].iloc[0]
            movement_ijk_dict[excel_key] = [i, j, k]
        else:
            from_4_approach = filter_direction(from_gdf, 'direction_4', approach_to_direction[approaches[0]])
            to_4_approach = filter_direction(to_gdf, 'direction_4', approaches[1])
            if len(from_4_approach) == 1 and len(to_4_approach) == 1:
                i = from_8_approach['ANode'].iloc[0]
                j = int_node
                k = from_8_approach['BNode'].iloc[0]
                movement_ijk_dict[excel_key] = [i, j, k]
                add_to_database = 2 # 2 result in displaying
            else:
                add_to_database = 0
                return add_to_database, movement_ijk_dict
    return add_to_database, movement_ijk_dict

