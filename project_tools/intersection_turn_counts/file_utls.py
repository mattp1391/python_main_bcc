import os
import sys
import glob
import pandas as pd
import csv


def check_file_path_is_folder_or_directory(file_path):
    file_type = None
    if os.path.isfile(file_path):
        file_type = 'file'
    elif os.path.isdir(file_path):
        file_type = 'directory'
    return file_type


def check_if_folder_exists(my_dir):
    folder_exists = os.path.isdir(my_dir)
    return folder_exists


def check_file_exists(my_file):
    file_exists = os.path.isfile(my_file)
    return file_exists


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


def create_csv_output_file(df, xl_file, output_folder=None, output_type=None):
    #print('create csv output')
    if output_type is None:
        output_type = '.csv.'
    output_name = xl_file.split('.xl', 1)[0] + output_type
    #print(output_name)
    #if output_folder is None:
    #    output_folder = r"D:\MP\projects\bcasm\log files\traffic_intersection_outputs"
    output_file = os.path.join(output_name, output_name)
    df.to_csv(output_file, index=False)
    return


def save_dataframe_log(df, filename):
    with open(filename, 'a', newline='') as f:
        df.to_csv(f, mode='a', header=f.tell() == 0, index=False, encoding="utf-8", quoting=csv.QUOTE_ALL)
    return


def exclude_files_already_assessed(all_files, assessed_file, col_check='file_name'):
    df_analysed = pd.read_csv(assessed_file, encoding='cp1252')
    df_file_names = df_analysed[col_check].unique().tolist()
    new_files = list(set(all_files) - set(df_file_names))
    return new_files


def make_clickable(url, name):
    return '<a href="{}" rel="noopener noreferrer" target="_blank">{}</a>'.format(url, name)


def get_file_size_in_bytes(file_path):
    """
    Get size of file at given path in bytes
    "
    Parameters
    ----------
    file_path

    Returns
    -------

    """
    size = os.path.getsize(file_path)
    return size

