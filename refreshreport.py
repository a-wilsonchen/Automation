from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium import webdriver
from utils import refresh_power_query
import time

options = webdriver.EdgeOptions()
options.add_argument("--disable-notifications")
driver = webdriver.Edge(options)
driver.get("https://portal.stratus.ms/spp-home")
time.sleep(10)
button = driver.find_element(
    By.XPATH, "//div[@id='cdk-accordion-child-0']/div/div/div/scc-app-card[2]/mat-card/button/div")
button.click()
time.sleep(10)
