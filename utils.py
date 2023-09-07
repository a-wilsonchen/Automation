# %% Loading Necessary Data
# TODO: Refactor,所有的Code都會需要dsmreader這個file
import pandas as pd
import pretty_errors
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
import os
import time
import win32com.client
import numpy as np
import pandas as pd
import datetime as dt
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


# refresh_power_query(join(T2_REPORT_DIRECTORY, "MinMax_2023-09-01.xlsx"))
# ? Automate Download DSM
# %%
"""
options = webdriver.ChromeOptions()
# Disable the automationcontrolled flag
options.add_argument("--disable-blink-features=AutomationControlled")
# Exclude the collection of eable-automation switches
options.add_experimental_option("excludeSwitches", ["enable-automation"])
# Turn-off userAutomationExtension
options.add_experimental_option("useAutomationExtension", False)
useragentarray = [
    "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
]

driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": useragentarray[0]})
"""
# options = webdriver.EdgeOptions()
# options.add_argument("--disable-notifications")
# driver = webdriver.Edge(options)
# driver.get("https://portal.stratus.ms/spp-home")
# time.sleep(10)
# button = driver.find_element(
#     By.XPATH, "//div[@id='cdk-accordion-child-0']/div/div/div/scc-app-card[2]/mat-card/button/div")
# button.click()
# time.sleep(10)
"""
search = driver.find_element(By.XPATH, "//input")
search.send_keys(ic_name)

button = driver.find_element(By.XPATH, "//div[@id='__next']/div[2]/div/div/div/div[2]/button/span[2]")
button.click()
# <input placeholder="" value="" class="jsx-4133345755 search-input">
actions = ActionChains(driver)
actions.send_keys(Keys.ENTER)
actions.perform()

# <button class=""><span class="jsx-4133345755 search-button-icon"><svg viewBox="0 0 512 512" class="jsx-1591298603"><path d="M505 442.7L405.3 343c-4.5-4.5-10.6-7-17-7H372c27.6-35.3 44-79.7 44-128C416 93.1 322.9 0 208 0S0 93.1 0 208s93.1 208 208 208c48.3 0 92.7-16.4 128-44v16.3c0 6.4 2.5 12.5 7 17l99.7 99.7c9.4 9.4 24.6 9.4 33.9 0l28.3-28.3c9.4-9.4 9.4-24.6.1-34zM208 336c-70.7 0-128-57.2-128-128 0-70.7 57.2-128 128-128 70.7 0 128 57.2 128 128 0 70.7-57.2 128-128 128z" class="jsx-1591298603"></path></svg></span><span class="jsx-4133345755 search-button-text">Search</span></button>
"""
# %%

# ? Automate Data Quality Checking Rules
# %% Read Data
product_info = singleprocessing_excel_file(
    join(T2_METRIC_DB, "DSM/Current Week"), sheet_name="1-Product Info")["1-Product Info"]
filename_latestfcst = os.listdir(join(T2_METRIC_DB, "T2 FCST/Current Week"))[0]
stratus_t2_fcst = pd.read_excel(join(T2_METRIC_DB, "T2 FCST/Current Week",
                                filename_latestfcst), sheet_name="Microsoft Forecast")

# %% #?Changing data type and remove duplicates.
product_info.rename(columns={"data_source": "Supplier",
                             "Description": "Part Subcategory"},
                    inplace=True
                    )
key_columns_product_info = ["Supplier", "MPN", "MSPN", "Part Subcategory", "Customer P/N"]
product_info = product_info[key_columns_product_info]

for col in product_info.columns:
    product_info[col] = product_info[col].astype(str)
    # print(f"Column {col} has {product_info[col].isna().sum()} null value, and {(product_info[col] == 'NaN').sum()} blank value")


stratus_t2_fcst = stratus_t2_fcst[["MFG Name", "MFG Part Number", "Microsoft Part Number", "SubCategory"]]
stratus_t2_fcst.rename(
    columns={
        "MFG Name": "Supplier",
        "MFG Part Number": "MPN",
        "Microsoft Part Number": "MSPN",
        "SubCategory": "Part Subcategory"},
    inplace=True
)
for col in stratus_t2_fcst.columns:
    stratus_t2_fcst[col] = stratus_t2_fcst[col].astype(str)
stratus_t2_fcst.drop_duplicates(inplace=True)

# ? AMPHENOL Special Treatment
product_info["MSPN"] = np.where(

    product_info["Supplier"] == "AMPHENOL",
    product_info["Customer P/N"],
    product_info["MSPN"]
)

# TODO:產出的LOG報表應該包含 Snapshot Date, Key, Supplier, Data Source, Type, Description,

# ? Product Info內部檢查

# %% #? 一個MSPN Mapped多個Part Subcategory
mspn_many_partsub = product_info.copy(deep=True)
mspn_many_partsub = mspn_many_partsub[["Supplier", "MSPN", "Part Subcategory"]]
mspn_many_partsub.drop_duplicates(inplace=True)
mspn_many_partsub["Count"] = mspn_many_partsub.groupby(["Supplier", "MSPN"]).transform('count')
mspn_many_partsub = mspn_many_partsub[mspn_many_partsub["Count"] >= 2]
mspn_many_partsub = mspn_many_partsub.groupby(["Supplier", "MSPN"]).agg(
    {"Part Subcategory": lambda x: ", ".join(x)}
).reset_index()

mspn_many_partsub["Data Source"] = "DSM, 1-Product Info"
mspn_many_partsub["Type"] = "One MSPN to Many Description"
mspn_many_partsub["Description"] = mspn_many_partsub["MSPN"] + " is mapped To " + mspn_many_partsub["Part Subcategory"]
mspn_many_partsub.drop(
    ["MSPN", "Part Subcategory"],
    axis=1,
    inplace=True
)

# %% #? 一個MPN Mapped到兩個MSPN
mpn_many_mspn = product_info.copy(deep=True)
mpn_many_mspn = mpn_many_mspn[["Supplier", "MSPN", "MPN"]]
mpn_many_mspn.drop_duplicates(inplace=True)
mpn_many_mspn["Count"] = mpn_many_mspn.groupby(["Supplier", "MPN"]).transform('count')
mpn_many_mspn = mpn_many_mspn[mpn_many_mspn["Count"] >= 2]
mpn_many_mspn = mpn_many_mspn.groupby(["Supplier", "MPN"]).agg(
    {"MSPN": lambda x: ", ".join(x)}
).reset_index()

mpn_many_mspn["Data Source"] = "DSM, 1-Product Info"
mpn_many_mspn["Type"] = "One MPN to Many MSPN"
mpn_many_mspn["Description"] = mpn_many_mspn["MPN"] + " is mapped To " + mpn_many_mspn["MSPN"]
mpn_many_mspn.drop(
    ["MPN", "MSPN"],
    axis=1,
    inplace=True
)
# %% #?和MSFT T2 FCST發布的數據比對, 看有沒有missing在Product Info裡面

missing_in_productInfo = stratus_t2_fcst.merge(
    product_info.drop("Customer P/N", axis=1),
    on=["Supplier", "Part Subcategory", "MSPN", "MPN"],
    how="outer",
    indicator=True,
    suffixes=("", "_ODM")
)

missing_in_productInfo = missing_in_productInfo[(missing_in_productInfo["_merge"] == "left_only")
                                                & (missing_in_productInfo["Supplier"].isin(SUPPLIER_LIST))]


missing_in_productInfo["Data Source"] = "DSM, 1-Product Info"
missing_in_productInfo["Type"] = "MPN/MSPN/Partsub Combination Missing"
missing_in_productInfo["Description"] = missing_in_productInfo["MPN"] + "/ " + \
    missing_in_productInfo["MSPN"] + "/ " + missing_in_productInfo["Part Subcategory"]
missing_in_productInfo.drop(
    ["MPN", "MSPN", "_merge", "Part Subcategory"],
    axis=1,
    inplace=True
)

# %%


# %%
