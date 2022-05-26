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

import os, sys
script_folder = r'C:\General\BCC_Software\Python\python_repository\python_library\python_main_bcc'
if script_folder not in sys.path:sys.path.append(script_folder)
from gis import osm_tools as osm













