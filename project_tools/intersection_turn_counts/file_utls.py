import os, sys, glob


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


