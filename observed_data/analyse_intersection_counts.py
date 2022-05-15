from win32com.client import Dispatch
from openpyxl.utils import get_column_letter, column_index_from_string
import numpy as np
import math
import geopandas as gpd
import pandas as pd
import numbers


# https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
shape_type_dict = {30: 'mso3DModel',
                   1: 'msoAutoShape',
                   2: 'msoCallout',
                   20: 'msoCanvas',
                   3: 'msoChart',
                   4: 'msoComment',
                   27: 'msoContentApp',
                   21: 'msoDiagram',
                   7: 'msoEmbeddedOLEObject',
                   8: 'msoFormControl',
                   5: 'msoFreeform',
                   28: 'msoGraphic',
                   6: 'msoGroup',
                   24: 'msoIgxGraphic',
                   22: 'msoInk',
                   23: 'msoInkComment',
                   9: 'msoLine',
                   31: 'msoLinked3DModel',
                   29: 'msoLinkedGraphic',
                   10: 'msoLinkedOLEObject',
                   11: 'msoLinkedPicture',
                   16: 'msoMedia',
                   12: 'msoOLEControlObject',
                   13: 'msoPicture',
                   14: 'msoPlaceholder',
                   18: 'msoScriptAnchor',
                   2: 'msoShapeTypeMixed',
                   25: 'msoSlicer',
                   19: 'msoTable',
                   17: 'msoTextBox',
                   15: 'msoTextEffect',
                   26: 'msoWebVideo'}

arrow_head_style = {1: 'msoArrowheadNone',
                    2: 'msoArrowheadTriangle',
                    3: 'msoArrowheadOpen',
                    4: 'msoArrowheadStealth',
                    5: 'msoArrowheadDiamond',
                    6: 'msoArrowheadOval',
                    -2: 'msoArrowheadStyleMixed'}

direction_dict = {0: 'N',
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



def variable_is_number(no):
    """
    Determines if a string is a number

    Parameters
    ----------
    no (string): any variable of type string

    Returns
    -------
    Bool: True if variable is number, False if it is not a number.
    """
    return isinstance(no, numbers.Number)


def find_turn_movement(worksheet, cell):
    """
    Finds the turn movement number and approach for a turn arrow object within an intersection turning survey
    Parameters
    ----------
    worksheet (excel worksheet object): worksheet object being assessed
    cell (row, col): cell value in which the arrow belongs

    Returns
    -------
    Turn movement number associated to the arrow and approach the turn belongs to
    """
    movement = None
    approach = None
    n_movement = worksheet.cells(cell[0] - 2, cell[1]).value
    e_movement = worksheet.cells(cell[0], cell[1] + 1).value
    s_movement = worksheet.cells(cell[0] + 2, cell[1]).value
    w_movement = worksheet.cells(cell[0], cell[1] - 1).value
    if variable_is_number(n_movement):
        movement = int(n_movement)
        approach = 'N'
    if variable_is_number(e_movement):
        movement = int(e_movement)
        approach = 'E'
    elif variable_is_number(s_movement):
        movement = int(s_movement)
        approach = 'S'
    elif variable_is_number(w_movement):
        movement = int(w_movement)
        approach = 'W'

    return movement, approach


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
    degrees = np.rad2deg((ang1 - ang2) % (2 * np.pi))
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
        destination_x = p2[1]
        origin_y = p1[0]
        destination_y = p2[0]
    else:
        origin_x = p1[0]
        destination_x = p2[0]
        origin_y = p1[1]
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


def find_point_positions(hor, ver, top, bottom, left, right):
    """
    Find the start and end poistions of an 'msoline' from MS excel
    Parameters
    ----------
    hor(int): represents the horizontal lip of the line.  0 if left to right, -1 if right to left.
    ver: represents the horizontal lip of the line.  0 if top to bottom, -1 if bottom to top.
    top: y position of the top point.
    bottom: y position of the bottom point
    left: x position of the left most point
    right x position of the right most point

    Returns
    -------

    """
    point_1 = (top, left)
    point_2 = (top, right)
    point_3 = (bottom, left)
    point_4 = (bottom, right)
    hor_vert_str = f"{hor}_{ver}"
    point_dict = {'0_0': [point_1, point_4, 'E_S'],
                  '0_-1': [point_3, point_2, 'N'],
                  '-1_0': [point_2, point_3, 'W'],
                  '-1_-1': [point_4, point_1, 'S']}
    line_pos = point_dict[hor_vert_str]
    return line_pos


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


def find_movement_dict_from_intersection_count(excel_file_path):
    """
    find relationship between turn movement numbers and origin to destination approach.  Approach is designated by North
    (N), East (E), South (S) or West (W).
    Parameters
    ----------
    excel_file_path (string): string of path to excel file

    Returns
    -------
    Dict: {movement_1: Origin_Destination, movement_2, origin_destination}
    """
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename=excel_file_path)
    ws = wb.Worksheets(1)
    intersection = ws.cells(4, 2).value
    turns_dict = {}
    peds_dict = {}
    text_dict = {}
    movement_dict = {}
    shapes = ws.shapes
    for sh in shapes:
        if shape_type_dict[sh.Type] == 'msoLine':
            line_name = sh.Name
            start_style = sh.Line.BeginArrowheadStyle
            end_style = sh.Line.EndArrowheadStyle
            cell_col = sh.TopLeftCell.Address.split("$")[1]
            cell_row = sh.TopLeftCell.Address.split("$")[2]
            cell = (int(cell_row), int(column_index_from_string(cell_col)))
            shape_top = sh.top
            shape_height = sh.height
            shape_left = sh.left
            shape_width = sh.width
            hor_flip = sh.HorizontalFlip
            ver_flip = sh.VerticalFlip
            line_pos = find_point_positions(hor_flip, ver_flip, shape_top, shape_top - shape_height, shape_left, shape_left + shape_width)
            angle = compass_angle(line_pos[0], line_pos[1], excel_cell_format=True)
            angle_round = int(custom_round(angle, base=45))
            if arrow_head_style[end_style] == 'msoArrowheadTriangle':
                movement, approach = find_turn_movement(ws, cell)
                direction = direction_dict[angle_round]
                movement_dict[movement] = f"{approach}_{direction}"
    wb.Close(True)
    return intersection, movement_dict


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


def add_angle_and_direction_columns(gdf, geometry_col='geometry', round_angle=45, angle_col='angle',
                                    angle_round_col='angle_round', direction_col='direction'):
    gdf
    links_to_count_gdf.loc[:, 'angle'] = links_to_count_gdf.apply(
        lambda row: compass_angle(row['AXco'], row['AYco'], row['BXco'], row['BYco']), axis=1)

