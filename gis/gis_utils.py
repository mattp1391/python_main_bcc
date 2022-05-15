from matplotlib import pyplot
from shapely.geometry import LineString


def parallel_offset_gpd(line_str, distance=5, side='left', resolution=16, join_style=1, mitre_limit=5.0):
    new_line_str = line_str.parallel_offset(distance, side, resolution=16, join_style=1, mitre_limit=5.0)
    return new_line_str

