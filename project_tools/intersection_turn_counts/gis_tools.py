from geopy import geocoders, Point
import matplotlib
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.colors import LinearSegmentedColormap
from PIL import Image
from geopy.distance import distance as geopy_distance
import contextily as cx
import time
from PIL import ImageGrab
from openpyxl.utils import get_column_letter, column_index_from_string
import numpy as np
import math
import geopandas as gpd
import pandas as pd
import numbers
from datetime import datetime, timedelta
import time
import pywintypes
from IPython.display import display
import os, sys
script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path:sys.path.append(script_folder)
from gis import osm_tools as osm



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


def find_angle_of_linestring(linestring, point_1_index=-1, point_2_index=0):
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
    ax = gdf_node_map.plot(markersize=marker_size, c='black', figsize=(10, 10))
    gdf_with_map.plot(colour_col, ax=ax, cmap=cmap, figsize=(10, 10), alpha=0.5, linewidth=10.0)

    for x, y, label, text_size in zip(gdf_node_map.geometry.x, gdf_node_map.geometry.y, gdf_node_map[label],
                                      gdf_node_map[text_size]):
        ax.annotate(label, xy=(x, y), xytext=(3, 3), textcoords="offset points", fontsize=text_size)
    ax.set_axis_off()
    #for x, y, label in zip(gdf_link['geometry'].x, gdf_link['geometry'].y, gdf_link['label']):
    #    ax.annotate(label, xy=(x, y), xytext=(3, 3), textcoords="offset points")
    #cx.add_basemap(ax, source=cx.providers.Esri.WorldImagery)
    cx.add_basemap(ax, source=cx.providers.HEREv3.satelliteDay(apiKey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ', proxies='165.225.226.22:10170'))
    cx.add_basemap(ax, source=cx.providers.HEREv3.mapLabels(apiKey='hROuZ5fSMweHJUgssiq6oehaPsd6u8-qMeF6CGN-SOQ', size=128, proxies='165.225.226.22:10170'))
    canvas = plt.get_current_fig_manager().canvas
    agg = canvas.switch_backends(FigureCanvasAgg)
    #agg.draw()
    s, (width, height) = agg.print_to_buffer()
    # Convert to a NumPy array.
    #X = np.frombuffer(s, np.uint8).reshape((height, width, 4))
    im = Image.frombytes("RGBA", (width, height), s)
    return im
