import pandas as pd
import numpy as np


def flatten_multi_index_columns_from_pivot(df, join_character='|',  level=None):
    """
    This function will flatten dataframes with multiple header rows.  It was designed to be used when a dataframe has
    been pivoted or grouped.

    -------
    Parameters df (pandas dataframe): pandas dataframe with multiIndex
    level (int): desired level for final headers.  If None, all headers will be used. 'top' will produce the top level, 'bottom' will use the bottom Integers may also
    be used to determine the level with 0 being the top and -1 the bottom.

    Returns
    -------
    dataframe with
    """
    #print('df_columns', df.columns.tolist())
    #columns = df.columns.map(lambda x: join_character.join([str(i) for i in x])).tolist()
    columns = df.columns.map(lambda x: x if isinstance(x, str) else join_character.join([str(i) for i in x])).tolist()
    #print('columns', columns)
    if level is not None:
        if isinstance(level, str):
            if level.lower() == 'top':
                level = 0
            elif level.lower() == 'bottom':
                level = -1
        new_columns = []
        for c in columns:
            new_columns.append(c.split("|")[level])
    else:
        new_columns = columns
    df.columns = new_columns
    df = df.reset_index()
    return df


def pivot_csv_file(file_name, values_col, pivot_col, index_cols, agg_func=None):
    df = pd.read_csv(file_name, encoding='cp1252')
    df_pivot = pd.pivot_table(df, values=values_col, columns=pivot_col,
                              index=index_cols)
                              # aggfunc={'count': np.mean})
    return df_pivot


def groupby_for_filtered_frame(df, filter_col, filter_val, group_by_cols):
    df = df[df[filter_col] == filter_val]
    df_grouped = df.groupby(group_by_cols).mean().reset_index()
    return df_grouped
