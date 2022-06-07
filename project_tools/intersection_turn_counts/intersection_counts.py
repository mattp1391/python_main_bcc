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
import csv

script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path: sys.path.append(script_folder)
from project_tools.intersection_turn_counts import gis_tools as gis
from project_tools.intersection_turn_counts import file_utls as fu

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

excel_border_checks = {'xlDiagonalDown': 5, 'xlDiagonalUp': 6, 'xlEdgeBottom': 9, 'xlEdgeLeft': 7, 'xlEdgeRight': 10,
                       'xlEdgeTop': 8, 'xlInsideHorizontal': 12, 'xlInsideVertical': 11}




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


def check_border(ws, row, col, location=9):
    border_style = ws.cells(row, col).Borders(location).LineStyle
    return border_style


def check_cell_value_strings(ws, row_from, search_strings, loop_rows=None):
    search_strings_lower = [str(x).lower() for x in search_strings]
    row_no_ = row_from
    output_row = -1
    if loop_rows is None:
        loop_rows = 99999
    while row_no_ <= row_from + loop_rows and output_row < 0:

        cell_value = ws.cells(row_no_, 1).value
        if cell_value is None:
            if search_strings is list:
                if None in search_strings:
                    output_row = row_no_
        elif isinstance(search_strings, list):
            # any(ext in url_string for ext in extensionsToCheck)
            if any(str(string_item).lower() in str(cell_value).lower() for string_item in search_strings_lower):
                output_row = row_no_
        else:
            if search_strings.lower() in str(cell_value).lower():
                output_row = row_no_
        row_no_ += 1
    return output_row


def create_header(ws, header_top, rows, col_left, columns):
    # ToDo:  May cause errors for peds if 3 teirs of data
    rows = [*range(header_top, header_top + rows + 1, 1)]
    cols = [*range(col_left, columns + 1, 1)]
    header_1 = []
    header_2 = []
    previous_head = None
    head_dict_for_row = {}
    for c in cols:
        top_row = True
        head = ""
        for r in rows:
            if top_row:
                cell_value = ws.cells(r, c).value
                if cell_value is not None:
                    header_1.append(str(cell_value))
                    previous_head = cell_value
                    top_row = False
                elif previous_head is not None:
                    header_1.append(previous_head)
                    top_row = False
            else:
                if ws.cells(r, c).Mergecells:
                    cell_value = ws.cells(r, c).MergeArea.Cells(1, 1).Value
                else:
                    cell_value = ws.cells(r, c).value
                if cell_value is not None:
                    if cell_value not in head and cell_value not in header_1:
                        if head == '':
                            head = cell_value
                        else:
                            head += f'_{cell_value}'
        head.replace('None|', '').replace('|None', '')
        header_2.append(head)
    head_array = [header_1, header_2]
    return head_array


def check_cell_has_border(ws, row_no_, col_no=1, location=8, next_row=1):
    border_style = -4142
    additional_rows = 0
    while border_style == -4142 and additional_rows < 10:
        border_style = check_border(ws, row_no_, col_no, location)  # 8 is location for top of cell refer top doc
        if border_style == -4142:
            additional_rows += next_row
            row_no_ = row_no_ + next_row
    return row_no_


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
            header_row = check_cell_has_border(ws, header_row, col_no=1, location=8, next_row=-1)

            row_no = header_row
            header_end = check_cell_value_strings(ws, row_no, search_strings=['(1/4 hr end)'], loop_rows=30)
            header_end = check_cell_has_border(ws, header_end, col_no=1, location=9, next_row=1)

            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xllinestyle?view=excel-pia


            use_columns = find_columns_used(ws, header_row, header_end)
            # display(use_columns)
            start_col = column_index_from_string(use_columns.split(':')[0])
            end_col = column_index_from_string(use_columns.split(':')[1])
            header = create_header(ws, header_row, header_end - header_row, start_col, end_col - start_col + 1)

            # header = create_header(ws, header_row, header_end - header_row, 1, header_end)
            end_df_row = check_cell_value_strings(ws, row_no, search_strings=df_end_strings, loop_rows=None) - 1
            if end_df_row == -1:
                search_for_additional_data = False
            row_no = end_df_row
            if search_for_additional_data:
                data_df = pd.read_excel(excel_file_path,
                                        sheet_name=sheet_name,
                                        skiprows=header_end,
                                        header=None,
                                        index_col=None,
                                        nrows=end_df_row - header_end,
                                        usecols=use_columns,
                                        parse_dates=False)
                pd.MultiIndex.from_arrays(header)
                if data_df is None:
                    return None
                else:
                    data_df.columns = pd.MultiIndex.from_arrays(header)
                    data_df.columns = data_df.columns.map(lambda x: '|'.join([str(i) for i in x]))
                    # display(1, data_df)
                    data_df = data_df.reset_index(drop=True)
                    #display(data_df.columns[0])
                    if data_df.columns[0].lower() == 'time|(1/4 hr end)':
                        data_df = data_df.rename(columns={data_df.columns[0]: 'TIME|(1/4 hr end)'})
                    df_melt = data_df.melt(id_vars=['TIME|(1/4 hr end)'], var_name='spreadsheet_movement|vehicle',
                                           value_name='count')
                    # display(2, df_melt)
                    df_melt[['temp_spreadsheet_movement', 'temp_vehicle']] = df_melt[
                        'spreadsheet_movement|vehicle'].str.split('|',
                                                                  expand=True)

                    df_melt['vehicle'] = np.where(
                        df_melt['temp_spreadsheet_movement'].str.lower().str.contains('pedestrian'),
                        'pedestrian', df_melt['temp_vehicle'])
                    df_melt['spreadsheet_movement'] = np.where(
                        df_melt['temp_spreadsheet_movement'].str.lower().str.contains('pedestrian'),
                        df_melt['temp_vehicle'], df_melt['temp_spreadsheet_movement'])

                    dataframes.append(df_melt)
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
    try:
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
    except:
        df = None
    return df


def find_columns_used(ws, top_row, bottom_row, col_from=1):
    col_no = col_from
    # if ws.cells(row_no, col_no).value is None:
    #    row_no += 1

    col_end_found = False
    while not col_end_found:
        row_no = top_row
        valid_cells = False
        while row_no <= bottom_row and not valid_cells:
            if ws.cells(row_no, col_no).value is not None:
                col_no += 1
                valid_cells = True

            else:
                if row_no == bottom_row:
                    col_end_found = True
                row_no += 1
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

def find_col_widths(ws, start_range=1, end_range=31):
    col_positions=[]
    current_width=0
    for col in range(start_range, end_range):
        cell_width = ws.cells(1, col).Width
        col_positions.append(current_width)
        current_width += cell_width
    return col_positions

'''
def find_nearest_cell():
    >> > from scipy import spatial
    >> > airports = [(10, 10), (20, 20), (30, 30), (40, 40)]
    >> > tree = spatial.KDTree(airports)
    >> > tree.query([(21, 21)])
    (array([1.41421356]), array([1]))
'''

def find_row_positions(ws, start_range=1, end_range=31):
    row_positions=[]
    row_hieght=0
    for row in range(start_range, end_range):
        cell_height = ws.cells(row, 1).Height
        row_positions.append(row_hieght)
        row_hieght += cell_height
    return row_positions


def find_cell_from_positions(row_pos, cell_pos, row_positions, col_positions):
    print('dod')


def find_nearest(list_, value):
    #print(array)
    array = np.asarray(list_)
    #print(2, array, value)
    if array.max() < value:
        cell_location = None
    else:
        min_value_greater = array[array > value].min()
        cell_location = list_.index(min_value_greater)
    return cell_location

def create_movement_dict(ws):
    movement_dict = {}
    shapes = ws.shapes
    '''
    text_box_dict = {}
    col_positions = find_col_widths(ws, start_range=1, end_range=31)
    #print(col_positions)
    row_positions = find_row_positions(ws, start_range=1, end_range=31)
    #print(row_positions)
    for sh in shapes:
        if shape_type_dict[sh.Type] == 'msoTextBox':
            shape_top = sh.top
            shape_left = sh.left
            #print('text box: ', sh.Name, sh.TopLeftCell.Address, sh.BottomRightCell.Address)
            pos_row = find_nearest(row_positions, shape_top)
            pos_col = find_nearest(col_positions, shape_left)
            text_box_value = str(ws.TextBoxes(sh.Name).Text)
            print('position: ', pos_row, pos_col, text_box_value, text_box_value.replace('.', '', 1).isdigit())
            #print(shape_left)
            if pos_row is not None and pos_col is not None and text_box_value.replace('.', '', 1).isdigit():
                text_box_dict[(pos_row, pos_col)] = text_box_value
    display(text_box_dict)
    '''
    for sh in shapes:
        if shape_type_dict[sh.Type] == 'msoLine':
            line_name = sh.Name
            start_style = sh.Line.BeginArrowheadStyle
            end_style = sh.Line.EndArrowheadStyle
            begin_style = sh.Line.BeginArrowheadStyle
            #print('begin_style', begin_style)
            cell_col = sh.TopLeftCell.Address.split("$")[1]
            cell_row = sh.TopLeftCell.Address.split("$")[2]
            cell = (int(cell_row), int(column_index_from_string(cell_col)))
            #print('cell: ' , sh.TopLeftCell.ColumnWidth)
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
            if (arrow_head_style[end_style] == 'msoArrowheadTriangle' and
                    arrow_head_style[begin_style] == 'msoArrowheadNone'):
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
    if df is None:
        wb.Close(True)
        return None, None
    survey_info_dict = get_survey_info(ws)
    intersection = survey_info_dict['survey_site']
    df_melt = df.drop(columns=['temp_spreadsheet_movement', 'temp_vehicle', 'spreadsheet_movement|vehicle'],
                      axis=1)
    df_melt['spreadsheet_movement'] = df_melt['spreadsheet_movement'].str.replace("movement ", "", case=False)
    df_melt['movement'] = df_melt['spreadsheet_movement'].map(movement_dict)
    df_melt['intersection'] = intersection
    df_melt['date'] = survey_info_dict['survey_date']
    df_melt['weather'] = survey_info_dict['survey_weather']
    df_melt = df_melt.rename(columns={'TIME|(1/4 hr end)': 'survey_time'})
    wb.Close(True)
    lat, lon, location = gis.geocode_coordinates(survey_info_dict['survey_site'], user_agent='Engineering_Services_BCC', api='here')
    if location is not None:
        location = location[0]
    df_melt['intersection_lat'] = lat
    df_melt['intersection_lon'] = lon
    #df_melt['geolocated_location'] = location
    df_melt['file_name'] = excel_file_path
    counts_gdf = gpd.GeoDataFrame(df_melt,
                                  geometry=gpd.points_from_xy(df_melt.intersection_lon, df_melt.intersection_lat),
                                  crs='epsg:4326')
    # df = pd.DataFrame.from_dict([movement_dict])
    return counts_gdf, movement_dict, intersection, location, lat, lon


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
    # ToDo: add capability for .xls files
    df = None
    movement_dict = None
    intersection = None
    location = None
    lat = None
    lon = None
    if not excel_file_path.endswith('.xls'):
        survey_sheet_info = find_survey_type(excel_file_path)
        survey_format = survey_sheet_info['survey_format']
        if survey_format is not None:
            df, movement_dict, intersection, location, lat, lon = survey_functions_map[survey_format](excel_file_path,
                                                                                  sheet_name=survey_sheet_info[
                                                                                      'sheet_name'])
            if movement_dict is not None:
                if 'None' in movement_dict:
                    # ToDo: update movement dict for text boxes??
                    # df = None
                    movement_dict = None
        else:
            df = None
            movement_dict = None

    return df, movement_dict, intersection, location, lat, lon


def find_survey_type(excel_file_path):
    wb = load_workbook(filename=excel_file_path)
    sheets = wb.sheetnames
    survey_format = None
    # ws.get_squared_range(min_col=1, min_row=1, max_col=1, max_row=10)
    if sheets == ['TABLE', 'excel_file_path']:
        survey_format = 'austraffic_1'
    elif wb[sheets[0]]["A1"].value is not None:
        if str(wb[sheets[0]]["A1"].value).lower() == 'austraffic video intersection count':
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


def check_text_boxes_for_movement():
    print('check_text_boxes')

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
    movement_found = False
    e_movement = worksheet.cells(cell[0], cell[1] + 1).value
    w_movement = worksheet.cells(cell[0], cell[1] - 1).value
    n_movement = worksheet.cells(cell[0] - 1, cell[1]).value
    if n_movement is None:
        if cell[0] > 2:
            n_movement = worksheet.cells(cell[0] - 2, cell[1]).value
    s_movement = worksheet.cells(cell[0] + 2, cell[1]).value
    if s_movement is None:
        s_movement = worksheet.cells(cell[0] + 1, cell[1]).value

    if variable_is_number(n_movement):
        movement = int(n_movement)
        approach = 'N'
    elif variable_is_number(e_movement):
        movement = int(e_movement)
        approach = 'E'
    elif variable_is_number(s_movement):
        movement = int(s_movement)
        approach = 'S'
    elif variable_is_number(w_movement):
        movement = int(w_movement)
        approach = 'W'
    #if movement is None:
    #    check_text_boxes_for_movement()
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


def check_log_file_exists_create_return_df(file_path, log_type=0):
    """
    Checks if log file exists.  If the file does not exist, a new log file is created.  The func
    Parameters
    ----------
    file_path (string): filepath as string

    Returns
    -------

    """

    path_type = fu.check_file_path_is_folder_or_directory(file_path)
    file_exists = False
    df = None
    if path_type == 'file':
        file_exists = True
    if not file_exists:
        with open(file_path, 'w') as f:
            if log_type == 0 or log_type.lower() == 'movement':
                f.write('"file","sheet","intersection","movement","i","j","k","comments"')
            elif log_type == 1 or log_type.lower() == 'data':
                f.write('"file","sheet","intersection","movement","i","j","k","comments"')
    return


def save_data_log(df, filename):
    with open(filename, 'a', newline='') as f:
        df.to_csv(f, mode='a', header=f.tell() == 0, index=False, quoting=csv.QUOTE_ALL)
    return


def save_movement_log(df, filename):
    with open(filename, 'a', newline='') as f:
        df.to_csv(f, mode='a', header=f.tell() == 0, index=False, quoting=csv.QUOTE_ALL)
    return


def find_files_assessed(df):
    files_assessed = df[df['file'].notnull()]
    return files_assessed
