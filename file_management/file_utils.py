import os
import shutil

import pandas as pd
import re
from typing import List, Optional


def create_folder(path):
    # type: (str) -> None
    file_string = re.compile(r"\..+$")
    if file_string.search(path) is not None:
        path = os.path.dirname(path)
    if not os.path.isdir(path):
        os.makedirs(path)