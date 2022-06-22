from geopy import geocoders, Point
from geopy.distance import distance as geopy_distance
import time
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

google_geo_code_key = 'AIzaSyDgFypRMb-gnE9eaFjiWjcdc6T4JpjGUAo'


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
