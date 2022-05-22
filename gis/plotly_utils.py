import plotly.express as px
import geopandas as gpd
import shapely.geometry
import numpy as np
from IPython.display import display

def plot_lines_from_gdf(geo_df, hover_name=None, mapbox_style=None, text=None, zoom=15):
    """

    Parameters
    ----------
    geo_df (geodataframe): geodataframe to be assessed
    hover_name (string): hover_string for map
    mapbox_style (string): 'open-street-map', 'white-bg', 'carto-positron', 'carto-darkmatter', 'stamen- terrain',
    'stamen-toner', 'stamen-watercolor'
    """
    if mapbox_style is None:
        mapbox_style = 'carto-darkmatter'
    lats = []
    lons = []
    hover_names = []
    texts = []
    geo_df = geo_df.to_crs(crs='epsg:4326')
    for feature, hover_name in zip(geo_df.geometry, geo_df[hover_name]):
        if isinstance(feature, shapely.geometry.linestring.LineString):
            linestrings = [feature]
        elif isinstance(feature, shapely.geometry.multilinestring.MultiLineString):
            linestrings = feature.geoms
        else:
            continue
        for linestring in linestrings:
            x, y = linestring.xy
            lats = np.append(lats, y)
            lons = np.append(lons, x)
            hover_names = np.append(hover_names, [hover_name] * len(y))
            texts = np.append(texts, [text] * len(y))
            lats = np.append(lats, None)
            lons = np.append(lons, None)
            hover_names = np.append(hover_names, None)
            texts = np.append(texts, None)
    fig = px.line_mapbox(lat=lats, lon=lons, hover_name=hover_names, labels=texts, text=texts,
                         mapbox_style=mapbox_style, zoom=zoom)
    fig.show()


def snip_excel_image():
    print('tbc')