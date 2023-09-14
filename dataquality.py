
# %% import pandas as pd
import pandas as pd
import os
import numpy as np
import pretty_errors
from os.path import join
from utils import singleprocessing_excel_file, T2_METRIC_DB, SUPPLIER_LIST, DSM_SHEETNAME
import utils

# ? Automate Data Quality Checking Rules
# %% Read Data

dsm_data = singleprocessing_excel_file(
    join(T2_METRIC_DB, "DSM/Current Week"), sheet_name=DSM_SHEETNAME)
filename_latestfcst = os.listdir(join(T2_METRIC_DB, "T2 FCST/Current Week"))[0]
stratus_t2_fcst = pd.read_excel(join(T2_METRIC_DB, "T2 FCST/Current Week",
                                filename_latestfcst), sheet_name="Microsoft Forecast")

# %% #?Changing data type and remove duplicates.
product_info = dsm_data["1-Product Info"].copy(deep=True)
product_info.rename(columns={"data_source": "Supplier",
                             "Description": "Part Subcategory"},
                    inplace=True
                    )
key_columns_product_info = ["Supplier", "MPN", "MSPN", "Part Subcategory", "Customer P/N"]
product_info = product_info[key_columns_product_info]

utils.clean_mapping_column(product_info, list(product_info.columns), inplace=True)
product_info = product_info[product_info.Supplier.isin(SUPPLIER_LIST)]

# ? AMPHENOL Special Treatment
product_info["MSPN"] = np.where(
    product_info["Supplier"] == "AMPHENOL",
    product_info["Customer P/N"],
    product_info["MSPN"]
)


stratus_t2_fcst = stratus_t2_fcst[["MFG Name", "MFG Part Number", "Microsoft Part Number", "SubCategory"]]
stratus_t2_fcst.rename(
    columns={
        "MFG Name": "Supplier",
        "MFG Part Number": "MPN",
        "Microsoft Part Number": "MSPN",
        "SubCategory": "Part Subcategory"},
    inplace=True
)
utils.clean_mapping_column(stratus_t2_fcst, list(stratus_t2_fcst.columns), inplace=True)
stratus_t2_fcst.drop_duplicates(inplace=True)
stratus_t2_fcst = stratus_t2_fcst[stratus_t2_fcst["Supplier"].isin(SUPPLIER_LIST)]


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

# MPN背對印到
mpn_mapped_to_different_mspn = stratus_t2_fcst.copy(deep=True)
benchmark = product_info.copy(deep=True)
mpn_mapped_to_different_mspn = mpn_mapped_to_different_mspn.merge(
    benchmark,
    on=["Supplier", "MPN"],
    how='left',
    suffixes=("", "_dsm")
)
mpn_mapped_to_different_mspn.dropna(subset=["MSPN_dsm"], inplace=True)
mpn_mapped_to_different_mspn = mpn_mapped_to_different_mspn[
    (mpn_mapped_to_different_mspn["MSPN"] != mpn_mapped_to_different_mspn["MSPN_dsm"]) &
    (mpn_mapped_to_different_mspn["MSPN_dsm"] != "nan")
]
mpn_mapped_to_different_mspn.drop_duplicates(subset=["Supplier", "MPN", "MSPN_dsm"], inplace=True)
mpn_mapped_to_different_mspn["Data Source"] = "DSM, 1-Product Info/ OnePDM"
mpn_mapped_to_different_mspn["Type"] = "MPN-MSPN Mapping Incconsistent "
mpn_mapped_to_different_mspn["Description"] = "MPN " + mpn_mapped_to_different_mspn["MPN"] + " mapped to " + \
    mpn_mapped_to_different_mspn["MSPN"] + " in Stratus, but mapped to " + \
    mpn_mapped_to_different_mspn["MSPN_dsm"] + " in product info tab."


# %% # TODO FC Summary
# ? 比對Description
fc_summary = dsm_data["FC Summary"].copy(deep=True)
utils.clean_mapping_column(fc_summary, ["Supplier", "MPN", "Description", "DSM_Date"], inplace=True)


# TODO 3-Finished Good
# ? 檢查MPN
# ? 檢查MPN和Description
# ? 檢查Location ID是否都是正確的
# ? 檢查In-transit Day是否是未來的?
# TODO 4-Productuin
# ? 檢查MPN
# ? 檢查MPN和Description
# ? 檢查Location ID是否都是正確的
# ? 檢查日期是否是未來的?
# TODO 6-Open Demand
# ? 檢查MPN
# ? 檢查MPN和Description
# ? 檢查T1 & T2 Location ID是否都是正確的
# ? 檢查日期是否是未來的?
# ? 檢查日期有沒有Null
# ? 檢查Transporation Method
# TODO 7-Shipped Demand
# ? 檢查前兩個月的貨物寄送有沒有變動
# ? 檢查PO
# ? 檢查Shiped From是不是正確的
# ? 檢查
# ? 檢查有沒有Shipment更動
# ? 檢查T1 & T2 Location ID
# ? 檢查日期有沒有未來

# %% #TODO 檢查IBP Data Quality
# ? 檢查PO數量
# ? 檢查T2 FCST (Ocean)
# ? 檢查T2 FCST Air和Ocean的總量
# ? 檢查T2 Ocean FCST的欄位有沒有
# %% #TODO 檢查FCST的品質
# ? 檢查WoW值的變化 + 前後周新的Part Sub和消失的Part Sub
# ? IBP, DSM數值差距(考量到FCST沒發布的情況下，IBP還是會做Netting這件事情, 可能很難檢核)
# ?

# %% #TODO 檢查Item Master
# %% #TODO 檢查TMC


# %%
