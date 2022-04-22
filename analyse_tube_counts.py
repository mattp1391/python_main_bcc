import pandas
import numpy

input_file = r"C:\General\BCC_Software\Python\python_repository\development_files\analyse_tube_counts\inputs\13759 " \
             r"Eastbound 23 Dobson Street ASCOT Between Racecourse Rd and Seymour Rd Class Volume 15 minute Report " \
             r".txt "

def find_tube_date_range(file_name):
    date_strings = None
    with open(file_name, "r") as f:
        for line in f:
            if "Filter time:" in line:
                date_strings = line.split('\t')[1]
                date_strings = date_strings.split(' (')[0]
                date_strings = date_strings.split(' => ')
                break
    return date_strings


def find_start_end_time(date_range):
    period_str = 'a'


dates = find_tube_date_range(input_file)
print(dates)