# %% Loading Necessary Data
# TODO: Refactor,所有的Code都會需要dsmreader這個file
import pandas as pd
import pretty_errors

import os
import time
import win32com.client
import pandas as pd
from os.path import join, isfile
from pathlib import Path
from multiprocessing import Pool
from typing import Union


# %%
BASE_DIR = Path(__file__).resolve().parent
DEBUG_MODE = True

DSM_SHEETNAME = [
    "FC Summary",
    "1-Product Info",
    "2-Factory Map",
    "3-Finished Goods",
    "4a-Production",
    "4b-Material Ready",
    "5-Open Demand",
    "7-Shipped Demand"]


# %%
USER_ID = os.getlogin()
T2_METRIC_DB = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database"
T2_MEASURES_DIRECTORY = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database/Measures"
T2_REPORT_DIRECTORY = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database/Report"
T2_MAPPING_DIRECTORY = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database/Mapping"
SUPPLIER_LIST = [
    "AAVID",
    "AMPHENOLCS",
    "ARTESYN",
    "ASSEMBLETECH",
    "AVC",
    "COOLERMASTER",
    "DELTA",
    "FCI",
    "FIT",
    "FLEX",
    "FURUKAWA",
    "GEIST",
    "INGRASYS",
    "JPC",
    "LENOVO",
    "LITEON",
    "LUXSHARE",
    "MOLEX",
    "NANJUENINTL",
    "NIDEC",
    "QUANTA",
    "RITTAL",
    "TE",
    "WELLTRUST",
    "WIWYNN",
    "ZT"
]


def time_function(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        output = func(*args, **kwargs)
        print(f"{func.__name__} takes {time.time() - start_time} seconds to finish.")
        return output

    return wrapper
# %%


def find_all_excel_predefined(path: str, keyword: str):
    def decorator(func):
        def wrapper(*args, **kwargs):
            func_globals = func.__globals__
            file = os.listdir(path)
            target_files_with_path = [join(path, f) for f in file if isfile(
                join(path, f)) and f.find(".xlsx") != -1 and f.find(keyword) != -1 and f.find("~") == -1]

            output = func(target_files_with_path, *args, **kwargs)

            return output
        return wrapper
    return decorator


def find_all_excel(path: str, keyword: Union[str, None] = None) -> list[str]:
    file = os.listdir(path)
    if keyword is not None:
        return [join(path, f) for f in file if isfile(
            join(path, f)) and f.find(".xlsx") != -1 and f.find(keyword) != -1 and f.find("~") == -1]
    else:
        return [join(path, f) for f in file if isfile(
            join(path, f)) and f.find(".xlsx") != -1 and f.find("~") == -1]
# %%


def read_single_dsm(file_path: str, sheet_name: Union[str, list[str]]) -> Union[pd.DataFrame, dict[str, pd.DataFrame]]:
    """Read single DSM file and get a single sheet in dataframe or get multiple sheets in dictionary

    Parameters
    ----------
    file_path : str
        The path where the target DSM file is located on local computer.
    sheet_name : Union[str, list[str]]
        A single sheet or multiple sheet. DSM valid sheet name include:
        FC Summary
        1-Product Info
        2-Factory Map
        3-Finished Goods
        4a-Production
        4b-Material Ready
        5-Open Demand
        7-Shipped Deman
    Returns
    -------
    Union[pd.DataFrame, dict[str, pd.DataFrame]]
        _description_
    """
    data = pd.read_excel(file_path, sheet_name=sheet_name)  # type: ignore

    if isinstance(sheet_name, list):

        for key, df in data.items():
            df.insert(loc=0, column='data_source',
                      value=file_path.split('\\')[-1].split('-')[1])
            df.insert(loc=1, column='DSM_Date', value='-'.join(file_path.split('\\')
                                                               [-1].split('-')[3:6]).replace('.xlsx', ''))
            df['DSM_Date'] = pd.to_datetime(df['DSM_Date'])

    else:
        data.insert(loc=0, column='data_source',
                    value=file_path.split('\\')[-1].split('-')[1])
        data.insert(loc=1, column='DSM_Date', value='-'.join(file_path.split('\\')
                    [-1].split('-')[3:6]).replace('.xlsx', ''))
        data['DSM_Date'] = pd.to_datetime(data['DSM_Date'])

    return data


@time_function
def singleprocessing_excel_file(file_path: str, sheet_name: Union[str, list[str]]) -> dict[str, pd.DataFrame]:

    if isinstance(sheet_name, str) & (sheet_name in DSM_SHEETNAME):
        file = find_all_excel(file_path, "DsmOutput-")

        output = []
        for f in file:
            output.append(read_single_dsm(f, sheet_name=sheet_name))

    elif all(elem in DSM_SHEETNAME for elem in sheet_name):
        file = find_all_excel(file_path, "DsmOutput-")

        output = []
        for f in file:
            output.append(read_single_dsm(f, sheet_name=sheet_name))

    else:
        raise ValueError(
            f"{sheet_name} contains a invalid sheet name or data type is incorrect.")

    # concatenate all reports.
    result_dict = {}
    if isinstance(sheet_name, list):
        for key in sheet_name:
            df_list = [d[key] for d in output]
            result_dict[key] = pd.concat(df_list, ignore_index=True)
    else:
        result_dict[sheet_name] = pd.concat(output, ignore_index=True)

    return result_dict


#! 若只需要讀一張DSM的話，效能不會比single CPU好
#! multiprocessing只能在__main__裡面執行，且debug的時候會導致data viewer打不開


@time_function
def multiprocessing_excel_file(file_path: str, sheet_name: Union[str, list]) -> dict[str, pd.DataFrame]:
    if isinstance(sheet_name, str) & (sheet_name in DSM_SHEETNAME):
        processors = Pool(4)
        file = [join(file_path, f) for f in os.listdir(file_path) if isfile(
            join(file_path, f)) and f.find("~") == -1 and f.find(".ini") == -1]

        inputs = [(f, sheet_name) for f in file]
        output = processors.starmap(read_single_dsm, inputs)
        processors.terminate()

    elif all(elem in DSM_SHEETNAME for elem in sheet_name):
        processors = Pool(4)
        file = [join(file_path, f) for f in os.listdir(file_path) if isfile(
            join(file_path, f)) and f.find("~") == -1]

        inputs = [(f, sheet_name) for f in file]
        output = processors.starmap(read_single_dsm, inputs)
        processors.terminate()

    else:
        raise ValueError(
            f"{sheet_name} contains a invalid sheet name or data type is incorrect.")

    # concatenate all reports.
    result_dict = {}
    if isinstance(sheet_name, list):
        for key in sheet_name:
            df_list = [d[key] for d in output]
            result_dict[key] = pd.concat(df_list, ignore_index=True)
    else:
        result_dict[sheet_name] = pd.concat(output, ignore_index=True)

    return result_dict


def clean_mapping_column(df: pd.DataFrame, columns: Union[str, list[str]], inplace: bool = False) -> Union[bool, pd.DataFrame]:
    """Ensure the columns you want to use to mapping is str type. Warning: null value will be converted to string "nan".

    Parameters
    ----------
    df : pd.DataFrame
        Input DataFrame
    columns : Union[str, list[str]]
        Single or mulitple target columns.
    inplace : bool, optional
        Pandas-like option, by default False

    Returns
    -------
    Union[bool, pd.DataFrame]
        Boolean if inplace equal true. 

    Raises
    ------
    ValueError
        If the columns not in input dataframe, the function will return which columns is not in input dataframe.

    """
    if isinstance(columns, str):
        if (columns not in list(df.columns)):
            raise ValueError(f"Columns {columns} not in DataFrame")
    else:
        if not all(ele in df.columns for ele in columns):
            raise ValueError(f"Columns {[col for col in columns if col not in df.columns]} not in DataFrame")

    if inplace:
        df[columns] = df[columns].astype(str)
        return True

    else:
        output_df = df.copy(deep=True)
        output_df[columns] = output_df[columns].astype(str)
        return output_df


def refresh_power_query(path: str) -> None:
    """Given an abosolute path to an excel file. The function will open up the excel and refresh power query.

    Parameters
    ----------
    path : str
        Absolute path to the excel to be refreshed.
    """
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(path)
    xl.Visible = True
    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()
    wb.save()
    xl.Quit()
