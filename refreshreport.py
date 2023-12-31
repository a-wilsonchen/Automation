
# %% Load Package & Set up edge
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os.path import join, isfile
from os import listdir
import utils
import time
import datetime as dt
import os
import shutil
import pandas as pd
import re
from multiprocessing import Pool

(Warning("Please manually download TMC file and refresh Minmax power query"))
# %%
USER_NAME = os.getlogin()
start_time = time.time()
today = dt.datetime.today()
monday_of_the_week = (today + dt.timedelta(days=- today.weekday())).strftime("%m%d")
monday_of_the_previous_week = (today + dt.timedelta(days=- today.weekday() - 7)).strftime("%m%d")
today_str = (today).strftime("%Y%m%d")
today_str_hyp = (today).strftime("%Y-%m-%d")
download_folder = f"C:/Users/{USER_NAME}/Downloads"
root_directory_dsm = f"{utils.T2_METRIC_DB}/DSM"
root_directory_dbs = f"{utils.T2_METRIC_DB}/DBS"
root_directory_measures = f"C:/Users/{USER_NAME}/OneDrive - Microsoft/General/T2 Metrix Database/Measures"

# %% #? Open up Edge Driver
options = webdriver.EdgeOptions()
options.add_argument("--disable-notifications")
driver = webdriver.Edge(options)
wait = WebDriverWait(driver, 30)

# %% #? Downlatest Forecast
driver.get("https://portal.stratus.ms/open-book?tab=publish-to-suppliers")
time.sleep(15)
driver.find_element(By.LINK_TEXT, "Tier2Final").click()

# %% Download DSM Data

for index, supplier in enumerate(utils.SUPPLIER_LIST):
    # driver.get(f"https://portal.stratus.ms/inventory-forecast-internal/company/{supplier}/all")
    if index == 0:
        driver.get(f"https://portal.stratus.ms/inventory-forecast-internal/company/{utils.SUPPLIER_LIST[0]}/all")
        time.sleep(20)
        # FIXME: FCI Exceptions: 當Joanne那邊確認FCINangton開始使用後移除這段程式碼
    elif supplier == "FCI":
        continue
    else:
        supplier_drop_down = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mat-select-0']/div/div[1]")))
        supplier_drop_down.click()
        supplier_button = wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[@ng-reflect-value = '{supplier}']")))
        supplier_button.click()
    time.sleep(3)
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-global-view/scc-sidebar-layout/mat-sidenav-container/mat-sidenav/div/scc-form[2]/form/scc-form-field[1]/div/div[2]/mat-form-field/div/div[1]/div"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[3]/div[2]/div/div/div/mat-option[2]/span"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/app-root/app-global-view/scc-sidebar-layout/mat-sidenav-container/mat-sidenav/div/scc-form[2]/form/scc-form-field[1]/div/div[2]/mat-form-field/div/div[1]/div"))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[3]/div[2]/div/div/div/mat-option[1]/span"))).click()

    download_button = driver.find_element(
        By.XPATH, "(//*[normalize-space(text()) and normalize-space(.)='DSM Analysis'])[1]/following::span[1]"
    )
    download_button.click()
    time.sleep(3)


# %% #?Download DBS Data
driver.get(f"https://s360.dbschenkerusa.com/main/portal")
flag = input("Please sign in your DBS account.(Press Y when you log in.)")
hub_list = [
    "MSCH",
    "VMI",
    "MSCZ",
    "MSCZ2"
]
report_list: dict[str, list[str]] = {
    "Inventory": ["Inventory Summary", "Item Master"],
    "Inbound": ["Inbound Summary"]
}

view_button_id = "ReportViewerMain_ctl08_ctl00"
img_button_id = "ReportViewerMain_ctl09_ctl04_ctl00_ButtonImg"
excel_button_link_text = "Excel"

if flag.upper() == "Y":
    for hub in hub_list:
        for report_section, reports in report_list.items():
            for report in reports:
                # driver.find_element(By.XPATH, "//li[@id='report-mega']/a/span").click()
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Reports"))).click()
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, hub))).click()
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, report_section))).click()
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, report))).click()
                time.sleep(1)
                wait.until(EC.frame_to_be_available_and_switch_to_it(0))
                wait.until(EC.frame_to_be_available_and_switch_to_it(0))
                wait.until(EC.element_to_be_clickable((By.ID, view_button_id))).click()
                wait.until(EC.element_to_be_clickable((By.ID, img_button_id))).click()
                time.sleep(1)
                current_window = driver.current_window_handle
                wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Excel"))).click()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[0])

print(f"The whole process take {(time.time() - start_time)//60 } mins and {(time.time() - start_time) % 60} secs")


# %% #? DSM: 創造新的資料夾，並將Current Week裡面的DSM移動到Archeve內， 將Download的資料移動到Current Week裡面


# %% #?將DBS Current Week裡面的Inventory Summary, Inbound Summary和Item Master的資料全部移動到舊的資料夾

old_dbs_oh_files = utils.find_all_excel(join(root_directory_dbs, "OH Current Week"), "InventorySummary.xlsx")
for file in old_dbs_oh_files:
    shutil.move(file, file.replace("OH Current Week", "Archived"))
old_dbs_ib_files = utils.find_all_excel(join(root_directory_dbs, "IB"), "InboundSummary.xlsx")
for file in old_dbs_ib_files:
    shutil.move(file, file.replace("IB", "Archived"))
old_dbs_item_master_files = utils.find_all_excel(join(root_directory_dbs, "ITEM MASTER"), "ItemMaster.xlsx")
for file in old_dbs_item_master_files:
    os.remove(file)

# %% #? 將Inventory Summary全部改名
new_dbs_oh_files = utils.find_all_excel(download_folder, "InventorySummary")
old_new_path_inventory = {}
for file in new_dbs_oh_files:
    with pd.ExcelFile(file) as xlsx:
        df = pd.read_excel(file, sheet_name="InventorySummary", nrows=1)
        hub_id = df.iloc[0, 0]
        match hub_id:
            case "JDAMSMX01":
                old_new_path_inventory[file] = join(download_folder, f"{today_str}_AMS_InventorySummary.xlsx")
            case "JDAMSCN01":
                old_new_path_inventory[file] = join(download_folder, f"{today_str}_APAC_InventorySummary.xlsx")
            case "JDAMSCZ01":
                old_new_path_inventory[file] = join(download_folder, f"{today_str}_EMEA_InventorySummary.xlsx")
            case "JDAMSCZ02":
                old_new_path_inventory[file] = join(download_folder, f"{today_str}_EMEA2_InventorySummary.xlsx")

for old_path, new_path in old_new_path_inventory.items():
    os.rename(old_path, new_path)


# %% #?將Item Master重新命名
old_new_path_item = {}
new_dbs_item_files = utils.find_all_excel(download_folder, "ItemMaster")
for file in new_dbs_item_files:
    with pd.ExcelFile(file) as xlsx:
        df = pd.read_excel(file, sheet_name="ItemMaster", nrows=1)
        hub_id = df.iloc[0, 0]
        match hub_id:
            case "JDAMSMX01":
                old_new_path_item[file] = join(download_folder, f"{today_str}_AMS_ItemMaster.xlsx")
            case "JDAMSCN01":
                old_new_path_item[file] = join(download_folder, f"{today_str}_APAC_ItemMaster.xlsx")
            case "JDAMSCZ01":
                old_new_path_item[file] = join(download_folder, f"{today_str}_EMEA_ItemMaster.xlsx")
            case "JDAMSCZ02":
                old_new_path_item[file] = join(download_folder, f"{today_str}_EMEA2_ItemMaster.xlsx")

for old_path, new_path in old_new_path_item.items():
    os.rename(old_path, new_path)


# %% #?將Inbound Summary重新命名
old_new_path_inbound = {}
new_dbs_item_files = utils.find_all_excel(download_folder, "InboundSummary")
for file in new_dbs_item_files:
    with pd.ExcelFile(file) as xlsx:
        df = pd.read_excel(file, sheet_name="InboundSummary", nrows=1)
        hub_id = df.iloc[0, 0]
        match hub_id:
            case "JDAMSMX01":
                old_new_path_inbound[file] = join(download_folder, f"{today_str}_AMS_InboundSummary.xlsx")
            case "JDAMSCN01":
                old_new_path_inbound[file] = join(download_folder, f"{today_str}_APAC_InboundSummary.xlsx")
            case "JDAMSCZ01":
                old_new_path_inbound[file] = join(download_folder, f"{today_str}_EMEA_InboundSummary.xlsx")
            case "JDAMSCZ02":
                old_new_path_inbound[file] = join(download_folder, f"{today_str}_EMEA2_InboundSummary.xlsx")

for old_path, new_path in old_new_path_inbound.items():
    os.rename(old_path, new_path)

# %% #?將DBS相關的檔案移動到最新的資料夾
for old_path in old_new_path_inventory.values():
    shutil.move(old_path, join(root_directory_dbs, "OH Current Week", old_path[old_path.find(today_str):]))
for old_path in old_new_path_inbound.values():
    shutil.move(old_path, join(root_directory_dbs, "IB", old_path[old_path.find(today_str):]))
for old_path in old_new_path_item.values():
    shutil.move(old_path, join(root_directory_dbs, "ITEM MASTER", old_path[old_path.find(today_str):]))

# ? 將DSM file從Current Week移動新創造的資料夾，並將Download Folder內移動到最新的資料夾
history_dsm_files = utils.find_all_excel(join(root_directory_dsm, "Current Week"), "DsmOutput-")
os.mkdir(join(root_directory_dsm, "Archived", monday_of_the_previous_week))
for file in history_dsm_files:
    shutil.move(file, join(root_directory_dsm, "Archived", monday_of_the_previous_week, file[file.find("DsmOutput-"):]))

current_week_files = utils.find_all_excel(download_folder, "DsmOutput-")
for file in current_week_files:
    shutil.move(file, join(root_directory_dsm, "Current Week", file[file.find("DsmOutput-"):]))


# %% #?複製全新的File(除了v2_MINMAX除外)

files_to_refresh = []
for folder in os.listdir(root_directory_measures):
    target_folder = join(root_directory_measures, folder)
    target_excel_files = utils.find_all_excel(target_folder)
    list_of_date = {dt.datetime.strptime(
        f.split("\\")[-1].split("_")[1].replace(".xlsx", ""), "%Y-%m-%d"): f for f in target_excel_files}
    last_week_file = list_of_date[max(list_of_date.keys())]
    this_week_file = re.sub("20[0-9][0-9]-[0-1][0-9]-[0-3][0-9]", today_str_hyp, last_week_file)
    shutil.copy(last_week_file, this_week_file)
    files_to_refresh.append(this_week_file)


# FIXME 目前所有的Power Query執行速度非常慢(一份檔案需要10分鐘)，看要切回Local Power Query還是要用Python
# print(Warning(""))
# category_updated_flag = input("Please manual refresh Part Subcategory Mapping Table file, and press Y after this.")
# if category_updated_flag.upper() == "Y":
#     for file in files_to_refresh:
#         utils.refresh_power_query(file)

