import os
import sqlite3
import datetime as dt
from contextlib import closing
from selenium import webdriver
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import time


# File locations & printer name
onedrive_location = os.path.join(os.getenv('USERPROFILE'), 'OneDrive', 'Shared Files - Shop', 'Python Scripts')
database_location = os.path.join(onedrive_location, 'Ginee', 'ginee_orders.db')
PRINTER = 'zk88t'

def setup_cursor():
    # Connects to db in autocommit mode
    conn = sqlite3.connect(database_location, isolation_level=None)
    cur = conn.cursor()
    return cur


def setup_driver(driver='Edge'):
    if driver == 'Chrome':
        driver_path = os.path.join(onedrive_location, 'chromedriver_win32', 'chromedriver.exe')
        driver = webdriver.Chrome(driver_path)
    else:
        driver_path = os.path.join(onedrive_location, 'edgedriver_win64', 'msedgedriver.exe')
        options = EdgeOptions()
        options.use_chromium = True
        # options.binary_location = PATH
        driver = Edge(driver_path, options = options)
    return driver


try:
    driver = setup_driver()
    driver.get('https://seller.ginee.com/')
    actions = ActionChains(driver)
    # Setting language to English
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'ant-select-arrow'))).click()
    languages = driver.find_elements_by_class_name("login-country-name")

    for language in languages:
        if language.text == "English":
            language.click()
            break

    # Logging in
    with closing(setup_cursor()) as cur:
        cur.execute("SELECT user, password FROM credentials WHERE platform = 'Ginee';")
        data = cur.fetchone()
        ginee_email, ginee_password = data[0], data[1]
    driver.find_element_by_xpath("//*[@placeholder='Please input your email']").send_keys(ginee_email)
    driver.find_element_by_xpath("//*[@placeholder='Please enter password']").send_keys(ginee_password)
    driver.find_element_by_tag_name("button").click()

    # Goes to order
    driver.implicitly_wait(10)
    driver.get("https://seller.ginee.com/main/order")

    # Switches frame & select "Paid" tab
    iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'myIframe')))
    driver.switch_to.frame(iframe)
    paid_tab = driver.find_element_by_xpath("//div[@aria-controls='rc-tabs-0-panel-PAID']")
    paid_tab.click()

    # Inserting pending data to sqlite
    next_page_exists = True
    with closing(setup_cursor()) as cur:
        while next_page_exists:     # Waits until table is fully loaded
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((
                                                By.XPATH, "//tbody[@class='ant-table-tbody']/tr")))
            time.sleep(5)
            table_rows = driver.find_elements_by_xpath("//tbody[@class='ant-table-tbody']/tr")
            for row in table_rows:
                ginee_order_id = row.get_attribute("data-row-key")
                order_number = row.text.split('\n')[0]
                store = 'Sookee' if 'Sookee Store' in row.text else 'Edge'
                print(f"INSERTING {store}  {order_number}  ({ginee_order_id})")
                cur.execute("""INSERT OR IGNORE INTO orders 
                                VALUES (?, ?, ?, ?);""", (ginee_order_id, order_number, dt.datetime.now(), store))
            next_page = driver.find_element_by_xpath("//li[@title='Next Page']")

            if next_page.get_attribute("aria-disabled") == 'false':
                next_page.click()
                print("CLICKING NEXT PAGE")
            else:
                next_page_exists = False
                print("END OF PAGE")

    # Goes to order page
    def get_order(order_number):
        with closing(setup_cursor()) as cur:
             cur.execute(f"SELECT ginee_order_id FROM orders WHERE order_number = '{order_number}';")
             ginee_order_id = cur.fetchone()[0]
        driver.get(f"https://seller.ginee.com/main/order/order-detail?orderId={ginee_order_id}")


    # Arrange Shipment (ready to ship)
    driver.find_element_by_xpath("//button[normalize-space()='Arrange Shipment']").click()
    driver.find_element_by_xpath("//button[normalize-space()='Arrange Shipment']").click()
    driver.find_element_by_xpath("//button[normalize-space()='Print Label']").click()

    # Print
    driver.execute_script('window.print();')    # Ctrl + P
    driver.switch_to.window(driver.window_handles[1])   # switches to print preview
    driver.find_element_by_id("selecttrigger-1").click()
    driver.find_element_by_xpath("//div[@title='OneNote for Windows 10']").click()
    driver.find_element_by_xpath("//button[normalize-space()='Print']").click()
finally:
    time.sleep(5)
    driver.quit()