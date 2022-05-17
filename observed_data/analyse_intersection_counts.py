from win32com.client import Dispatch
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import numpy as np
import math
import geopandas as gpd
import pandas as pd
import numbers
from datetime import datetime, timedelta
import time
import pywintypes



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
                   -2: 'msoShapeTypeMixed',
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

# https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
excel_horizontal_alignment_dict = {-4108: 'Center',
                                   7: 'Center across selection',
                                   -4117: 'Distribute',
                                   5: 'Fill',
                                   1: 'Align according to data type',
                                   -4130: 'Justify',
                                   -4131: 'Left',
                                   -4152: 'Right'}




def ole_to_date_str(pywin_datetime):
    py_datetime = datetime(
                    year=pywin_datetime.year,
                    month=pywin_datetime.month,
                    day=pywin_datetime.day,
                    hour=pywin_datetime.hour,
                    minute=pywin_datetime.minute,
                    second=pywin_datetime.second)
    date_str = py_datetime.strftime("%d/%m/%Y")
    return date_str


def find_number_of_columns(ws, row_no, col_no):
    while ws.cells(row_no, col_no).value is None:
        col_no += 1
    return col_no


def find_count_data_rows(ws, row_from, loop_rows, search_strings):
    row_no = row_from
    row_no_output = -1
    while row_no <= row_from + loop_rows and row_no_output < 0:
        cell_value = ws.cells(row_no, 1).value
        if isinstance(search_strings, list):
            if any(string_item.lower() in str(cell_value).lower(0) for string_item in search_strings):
                row_no_output = row_no
        else:
            if search_strings.lower() in str(cell_value).lower():
                row_no_output = row_no
        row_no += 1
    return row_no_output


def get_data(ws):
    row_loop = range(1, 11)
    col_loop = range(1, 11)
    survey_info_dict = {'survey_site': None, 'survey_date': None, 'survey_weather': None}
    survey_date = None
    survey_site = None
    survey_weather = None
    for col in col_loop:
        for row in row_loop:
            cell_value = ws.cells(row, col).value
            if 'location' in str(cell_value).lower():
                survey_info_dict['survey_site'] = ws.cells(row, col + 1).value
            elif 'date' in str(cell_value).lower():
                pywin_datetime = ws.cells(row, col + 1).value
                #print(type(ole_date))
                py_date = ole_to_date_str(pywin_datetime)
                survey_info_dict['survey_date'] = py_date
            elif 'weather' in str(cell_value).lower():
                survey_info_dict['survey_weather'] = ws.cells(row, col + 1).value

    row_start = 1

    return survey_info_dict


def pandas_read_excel_multi_index_with_use_cols(excel_file, sheet_name=0, header_from=0, header_to=0, n_rows=None,
                                                use_cols=None):
    """
    Pandas is unable to specify usecols when extracting multi-index dataframe from excel.  This function
    provides a workaround for this feature.
    Parameters
    ----------
    excel_file (string): Path to excel file to be read
    sheet_name (string or int): name (string) or index (int) of excel sheet to be read
    header_from (int): first line of haeder.
    header_to (int): last line of header
    n_rows (int): number of rows to be included in dataframe
    use_cols (string or int): rows to be used in dataframe

    Returns
    -------
    object (dataframe)
    """
    #df = pd.read_excel(excel_file, sheet_name=sheet_name, header=[header_from, header_to], nrows=n_rows, usecols=use_cols)

    df = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       header=header_to+1,
                       index_col=[header_from, header_to],
                       nrows=10,
                       usecols=use_cols,
                       parse_dates=False)
    index = pd.read_excel(excel_file,
                          sheet_name=sheet_name,
                          header=None,
                          skiprows=header_from,
                          index_col=[header_from, header_to],
                          nrows=2,
                          usecols=use_cols,
                          parse_dates=False)
    index = index.fillna(method='ffill', axis=1)
    df.columns = pd.MultiIndex.from_arrays(index.values)
    df.columns = df.columns.map(lambda x: '|'.join([str(i) for i in x]))
    df = df.reset_index(drop=True)
    return df


def find_column_count(ws, row_no, col_from=1):
    col_no = col_from
    while ws.cells(row_no, col_no).value is not None:
        col_no += 1
    use_columns = list(range(col_from-1, col_no))
    return use_columns


def convert_spreadsheet_to_dataframe(ws, search_strings_headings, search_strings_end):
    print('test')


def get_austraffic_1_survey_data(excel_file_path, sheet_name):
    # ToDo: updated doc string
    """
    find relationship between turn movement numbers and origin to destination approach.  Approach is designated by North
    (N), East (E), South (S) or West (W).
    Parameters
    ----------
    excel_file_path (string): string of path to excel file

    Returns
    -------
    DataFrame: {movement_1: Origin_Destination, movement_2, origin_destination}
    """
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename=excel_file_path)
    ws = wb.Worksheets(1)
    get_data(ws)
    movement_dict = {}
    intersection = ws.cells(4, 2).value
    #movement_dict['intersection'] = intersection
    turns_dict = {}
    peds_dict = {}
    text_dict = {}
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
            line_pos = find_point_positions(hor_flip, ver_flip, shape_top, shape_top - shape_height, shape_left,
                                            shape_left + shape_width)
            angle = compass_angle(line_pos[0], line_pos[1], excel_cell_format=True)
            angle_round = int(custom_round(angle, base=45))
            if arrow_head_style[end_style] == 'msoArrowheadTriangle':
                movement, approach = find_turn_movement(ws, cell)
                direction = direction_dict[angle_round]
                movement_dict[str(movement)] = f"{approach}_{direction}"

    row_no = find_count_data_rows(ws, row_from=1, loop_rows=30, search_strings='Time')
    use_columns = find_column_count(ws, row_no + 1)
    data_df = pandas_read_excel_multi_index_with_use_cols(excel_file_path, sheet_name, header_from=row_no - 1,
                                                         header_to=row_no, n_rows=5,
                                                         use_cols=use_columns)
    #data_df = add_survey_details_to_dataframe(data_df, )
    wb.Close(True)
    print(movement_dict)
    df = pd.DataFrame.from_dict([movement_dict])
    return data_df


def get_ttm_1_survey_data(ws):
    df = pd.DataFrame
    return df


def get_matrix_1_survey_data(ws):
    df = pd.DataFrame
    return df


def get_data_audit_systems_1_survey_data(ws):
    df = pd.DataFrame
    return df


survey_functions_map = {"austraffic_1": get_austraffic_1_survey_data, "ttm_1" : get_ttm_1_survey_data,
                        "matrix_1": get_matrix_1_survey_data,
                        " data_audit_systems_1": get_data_audit_systems_1_survey_data}


def get_survey_data_main(excel_file_path, survey_format, sheet_name):
    survey_sheet_info = find_survey_type(excel_file_path)
    survey_format = survey_sheet_info['survey_format']
    df = survey_functions_map[survey_format](excel_file_path, sheet_name=survey_sheet_info['sheet_name'])
    return df


def find_survey_type(excel_file_path):
    wb = load_workbook(filename=excel_file_path)
    sheets = wb.sheetnames
    survey_format = None
    # ws.get_squared_range(min_col=1, min_row=1, max_col=1, max_row=10)
    if sheets == ['TABLE', 'excel_file_path']:
        survey_format = 'austraffic_1'
    elif wb[sheets[0]]["A1"].value.lower() == 'austraffic video intersection count':
        survey_format = 'austraffic_1'
    sheet_name = sheets[0]
    survey_sheet_info = {'sheet_name': sheet_name,
                         'survey_format': survey_format}
    return survey_sheet_info


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
    gdf.loc[:, 'direction_4'] = gdf['angle_round_90'].map(direction_dict)
    gdf.loc[:, 'direction_8'] = gdf['angle_round_45'].map(direction_dict)
    return gdf
