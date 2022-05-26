from win32com.client import Dispatch
from openpyxl import load_workbook
from PIL import ImageGrab
from openpyxl.utils import get_column_letter, column_index_from_string
import geopandas as gpd
import pandas as pd
import numpy as np
import numbers
from datetime import datetime
from IPython.display import display
import sys

script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path: sys.path.append(script_folder)
from project_tools.intersection_turn_counts import gis_tools as gis

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


def check_cell_value_strings(ws, row_from, search_strings, loop_rows=None):
    search_strings_lower = [str(x).lower() for x in search_strings]
    row_no = row_from
    output_row = -1
    if loop_rows is None:
        loop_rows = 99999
    while row_no <= row_from + loop_rows and output_row < 0:
        cell_value = ws.cells(row_no, 1).value
        if cell_value is None:
            if search_strings is list:
                if None in search_strings:
                    output_row = row_no
        elif isinstance(search_strings, list):
            # any(ext in url_string for ext in extensionsToCheck)
            if any(str(string_item).lower() in str(cell_value).lower() for string_item in search_strings_lower):
                output_row = row_no
        else:
            if search_strings.lower() in str(cell_value).lower():
                output_row = row_no
        row_no += 1
    return output_row


def find_count_data_rows(excel_file_path, sheet_name, ws, row_from, header_strings, df_end_strings):
    search_for_additional_data = True
    spreadsheet_df = None
    dataframes = []
    row_no = row_from
    while search_for_additional_data:
        header_row = check_cell_value_strings(ws, row_no, search_strings=header_strings, loop_rows=30)
        if header_row == -1:
            search_for_additional_data = False
        else:
            row_no = header_row
            use_columns = find_columns_used(ws, header_row + 1)
            end_df_row = check_cell_value_strings(ws, row_no, search_strings=df_end_strings, loop_rows=None) - 1
            if end_df_row == -1:
                search_for_additional_data = False
            row_no = end_df_row
            if search_for_additional_data:
                data_df = pandas_read_excel_multi_index_with_use_cols(excel_file_path, sheet_name=sheet_name,
                                                                      header_from=header_row - 1, header_to=header_row,
                                                                      n_rows=end_df_row - (header_row + 1),
                                                                      use_cols=use_columns)
                dataframes.append(data_df)
    if dataframes:
        spreadsheet_df = pd.concat(dataframes)
    return spreadsheet_df


def get_survey_info(ws):
    row_loop = range(1, 11)
    col_loop = range(1, 11)
    survey_info_dict = {'survey_site': None, 'survey_date': None, 'survey_weather': None}
    for col in col_loop:
        for row in row_loop:
            cell_value = ws.cells(row, col).value
            if 'location' in str(cell_value).lower():
                survey_info_dict['survey_site'] = ws.cells(row, col + 1).value
            elif 'date' in str(cell_value).lower():
                pywin_datetime = ws.cells(row, col + 1).value
                py_date = ole_to_date_str(pywin_datetime)
                survey_info_dict['survey_date'] = py_date
            elif 'weather' in str(cell_value).lower():
                survey_info_dict['survey_weather'] = ws.cells(row, col + 1).value

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
    # df = pd.read_excel(excel_file, sheet_name=sheet_name, header=[header_from, header_to], nrows=n_rows, usecols=use_cols
    df = pd.read_excel(excel_file,
                       sheet_name=sheet_name,
                       header=header_to,
                       index_col=[header_from, header_to],
                       nrows=n_rows,
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


def find_columns_used(ws, row_no, col_from=1):
    col_no = col_from
    while ws.cells(row_no, col_no).value is not None:
        col_no += 1
    start_col = get_column_letter(col_from)
    end_col = get_column_letter(col_no - 1)

    use_columns = f'{start_col}:{end_col}'
    return use_columns


'''
def convert_spreadsheet_to_dataframe(excel_file_path, ws, sheet_name=0, search_strings_headings='Time',
                                     search_strings_end=[None, 'Peak', 'Total'], row_no=1, loop_rows=30):
    # ToDo: check if this is used

    while row_no != -1:
        row_details = find_count_data_rows(excel_file_path, sheet_name, ws, row_from=1, loop_rows=30,
                                           header_strings=search_strings_headings, df_end_strings=search_strings_end)

        data_df = pandas_read_excel_multi_index_with_use_cols(excel_file_path, sheet_name, header_from=row_no - 1,
                                                              header_to=row_no, n_rows=5, use_cols=use_columns)
        row_no = row_details[-1]
'''


def create_movement_dict(ws):
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
            line_pos = find_point_positions(hor_flip, ver_flip, shape_top, shape_top - shape_height, shape_left,
                                            shape_left + shape_width)
            angle = gis.compass_angle(line_pos[0], line_pos[1], excel_cell_format=True)
            angle_round = int(gis.custom_round(angle, base=45))
            if arrow_head_style[end_style] == 'msoArrowheadTriangle':
                movement, approach = find_turn_movement(ws, cell)
                direction_dict = gis.get_direction_dict()
                direction = direction_dict[angle_round]
                movement_dict[str(movement)] = f"{approach}_{direction}"
    return movement_dict


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
    ws = wb.Worksheets(sheet_name)
    movement_dict = create_movement_dict(ws)
    df = find_count_data_rows(excel_file_path, sheet_name, ws, row_from=1, header_strings='Time',
                              df_end_strings=[None, 'Peak', 'Total'])
    survey_info_dict = get_survey_info(ws)
    df_melt = df.melt(id_vars=['TIME|(1/4 hr end)'], var_name='spreadsheet_movement|vehicle', value_name='count')
    df_melt[['temp_spreadsheet_movement', 'temp_vehicle']] = df_melt['spreadsheet_movement|vehicle'].str.split('|',
                                                                                                               expand=True)
    df_melt['vehicle'] = np.where(df_melt['temp_spreadsheet_movement'].str.lower().str.contains('pedestrian'),
                                  'pedestrian', df_melt['temp_vehicle'])
    df_melt['spreadsheet_movement'] = np.where(df_melt['temp_spreadsheet_movement'].str.lower().str.contains('pedestrian'),
                                               df_melt['temp_vehicle'], df_melt['temp_spreadsheet_movement'])
    df_melt = df_melt.drop(columns=['temp_spreadsheet_movement', 'temp_vehicle', 'spreadsheet_movement|vehicle'],
                           axis=1)
    df_melt['spreadsheet_movement'] = df_melt['spreadsheet_movement'].str.replace("movement ", "", case=False)
    df_melt['movement'] = df_melt['spreadsheet_movement'].map(movement_dict)
    df_melt['intersection'] = survey_info_dict['survey_site']
    df_melt['date'] = survey_info_dict['survey_date']
    df_melt['weather'] = survey_info_dict['survey_weather']
    df_melt = df_melt.rename(columns={'TIME|(1/4 hr end)': 'survey_time'})
    wb.Close(True)
    coords = gis.geocode_coordinates(survey_info_dict['survey_site'], user_agent='Engineering_Services_BCC', api='here')
    df_melt['intersection_lat'] = coords[0]
    df_melt['intersection_lon'] = coords[1]
    counts_gdf = gpd.GeoDataFrame(df_melt,
                                  geometry=gpd.points_from_xy(df_melt.intersection_lon, df_melt.intersection_lat),
                                  crs='epsg:4326')

    # df = pd.DataFrame.from_dict([movement_dict])
    return counts_gdf, movement_dict


def get_ttm_1_survey_data(ws):
    df = pd.DataFrame
    return df


def get_matrix_1_survey_data(ws):
    df = pd.DataFrame
    return df


def get_data_audit_systems_1_survey_data(ws):
    df = pd.DataFrame
    return df


survey_functions_map = {"austraffic_1": get_austraffic_1_survey_data, "ttm_1": get_ttm_1_survey_data,
                        "matrix_1": get_matrix_1_survey_data,
                        " data_audit_systems_1": get_data_audit_systems_1_survey_data}


def get_survey_data_main(excel_file_path):
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
    wb.close()
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


def find_point_positions(hor, ver, top, bottom, left, right):
    """
    Find the start and end positions of an 'msoline' from MS excel
    Parameters
    ----------
    hor: represents the horizontal lip of the line.  0 if left to right, -1 if right to left.
    ver: represents the horizontal lip of the line.  0 if top to bottom, -1 if bottom to top.
    top: y position of the top point.
    bottom: y position of the bottom point
    left: x position of the left most point
    right: x position of the right most point

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


def get_excel_image(file_path, sheet_name, range=None):
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(file_path)
    ws = wb.Worksheets[sheet_name]
    ws.Range(ws.Cells(1, 8), ws.Cells(15, 36)).Copy()
    img = ImageGrab.grabclipboard()
    wb.Application.CutCopyMode = False
    wb.Close(True)
    return img


def check_approach_match(xl_approaches, gis_approaches_4, gis_approaches_8=None):
    print('todo')


