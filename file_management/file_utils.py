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
    return


def replace_root(path, root, new_root):
    # type: (str, str, str) -> str

    rel = os.path.relpath(path, root)
    return os.path.join(new_root, rel)