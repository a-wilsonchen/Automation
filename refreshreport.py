
# %% Load Package & Set up edge
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import refresh_power_query, SUPPLIER_LIST
import time
# %%

options = webdriver.EdgeOptions()
options.add_argument("--disable-notifications")
driver = webdriver.Edge(options)
wait = WebDriverWait(driver, 120)
# driver.get("https://portal.stratus.ms/spp-home")
# time.sleep(10)

# %% Download DSM
for index, supplier in enumerate(SUPPLIER_LIST):
    driver.get(f"https://portal.stratus.ms/inventory-forecast-internal/company/{supplier}/all")
    if (index == 0):
        time.sleep(20)
    else:
        time.sleep(15)

    button = driver.find_element(
        By.XPATH, "(//*[normalize-space(text()) and normalize-space(.)='DSM Analysis'])[1]/following::span[1]"
    )
    button.click()
    time.sleep(3)
# button = driver.find_element(
#     By.XPATH, "//div[@id='cdk-accordion-child-0']/div/div/div/scc-app-card[2]/mat-card/button/div")
# button.click()
# time.sleep(10)

# %% Download DBS Data
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
