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