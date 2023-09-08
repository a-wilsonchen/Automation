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
    "AMPHENOLCS",
    "ARTESYN",
    "ASSEMBLETECH",
    "AVC",
    "COOLERMASTER",
    "DELTA",
    "FIT",
    "FLEX",
    "FURUKAWA",
    "GEIST",
    "HARTING",
    "INGRASYS",
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
        file = [join(file_path, f) for f in os.listdir(file_path) if isfile(
            join(file_path, f)) and f.find("~") == -1 and f.find(".ini") == -1]

        output = []
        for f in file:
            output.append(read_single_dsm(f, sheet_name=sheet_name))

    elif all(elem in DSM_SHEETNAME for elem in sheet_name):
        file = [join(file_path, f) for f in os.listdir(file_path) if (isfile(
            join(file_path, f))) and f.find("~") == -1 and (f.find("DsmOutput") != -1)]

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


# TODO: Need to revise the code based on below method. https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
@time_function
def refresh_power_query(path):
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(path)
    xl.Visible = True
    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()
    wb.save()
    xl.Quit()
