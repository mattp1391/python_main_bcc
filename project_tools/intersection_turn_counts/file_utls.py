from win32com.client import Dispatch
from openpyxl import load_workbook, Workbook
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
import pathlib
import os, sys, glob

script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path: sys.path.append(script_folder)


def check_if_folder_exists(my_dir):
    folder_exists = os.path.isdir(my_dir)
    return folder_exists


def create_folder_if_does_not_exist(my_dir):
    folder_exists = os.path.isdir(my_dir)
    if not folder_exists:
        os.makedirs(my_dir)
        print("created folder : ", my_dir)
    else:
        print(my_dir, "folder already exists.")


def get_list_of_files_in_directory(path, file_type = None, sub_folders=False):
    if file_type is None:
        file_type = ".*"
    if sub_folders:
        text_files = glob.glob(f"{path}/**/*{file_type}", recursive=True)
    else:
        text_files = glob.glob(f"{path}/*{file_type}", recursive=True)
    return text_files



