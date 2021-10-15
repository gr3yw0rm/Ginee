import os
import time
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
from selenium.common.exceptions import WebDriverException, StaleElementReferenceException, ElementClickInterceptedException
from tkinter import *
from tkinter import ttk
from threading import Thread, current_thread


# File locations & printer name
onedrive_location = os.path.join(os.getenv('USERPROFILE'), 'OneDrive', 'Shared Files - Shop', 'Python Scripts')
database_location = os.path.join(onedrive_location, 'Ginee', 'ginee_orders.db')
PRINTER = 'ZDesigner GK888t'
PRINTER = 'Microsoft Print to PDF'

def setup_cursor():
    # Connects to db in autocommit mode
    conn = sqlite3.connect(database_location, isolation_level=None)
    cur = conn.cursor()
    return cur


def setup_driver(driver='Edge', headless=False, maximized=False, zoom_level=1.0, window_position=(0, 0)):
    if driver == 'Chrome':
        driver_path = os.path.join(onedrive_location, 'chromedriver_win32', 'chromedriver.exe')
        driver = webdriver.Chrome(driver_path)
    else:
        driver_path = os.path.join(onedrive_location, 'edgedriver_win64', 'msedgedriver.exe')
        options = EdgeOptions()
        options.use_chromium = True
        if headless:
            options.add_argument('headless')
            options.add_argument('disable-gpu')
        # options.binary_location = PATH
        driver = Edge(driver_path, options = options)
    if zoom_level != 1.0:
        driver.get('edge://settings/')
        driver.execute_script(f'chrome.settingsPrivate.setDefaultZoom({zoom_level});')
    if window_position != (0, 0):
        driver.set_window_position(window_position[0], window_position[1])
    if maximized:
        driver.maximize_window()
    actions = ActionChains(driver)
    return driver


def login(driver):
    print("LOGGING IN TO GINEE")
    # Goes to website
    # driver.implicitly_wait(5)
    driver.get('https://seller.ginee.com/')

    with closing(setup_cursor()) as cur:
        cur.execute("SELECT user, password FROM credentials WHERE platform = 'Ginee';")
        data = cur.fetchone()
        ginee_email, ginee_password = data[0], data[1]

    while driver.find_elements_by_xpath('//button[normalize-space()="Login"]'):     #> BUG! Not logging in the first time
        # Setting language to English
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'ant-select-arrow'))).click()
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//li[normalize-space()="English"]'))).click()
        # Logging in
        driver.find_element_by_xpath("//*[@placeholder='Please input your email']").send_keys(ginee_email)
        driver.find_element_by_xpath("//*[@placeholder='Please enter password']").send_keys(ginee_password)
        driver.find_element_by_xpath('//button[normalize-space()="Login"]').click()
        time.sleep(3)
    print("\tSucessfully logged in.")


def scrape(driver=None, headless=False):
    print("SCRAPING GINEE ORDER IDs")
    if headless and driver is None:
        driver = setup_driver(headless=True)
        login(driver)

    # Goes to order
    driver.implicitly_wait(10)
    driver.get("https://seller.ginee.com/main/order")

    # Switches frame & select "Paid" tab
    iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'myIframe')))
    driver.switch_to.frame(iframe)
    paid_tab = driver.find_element_by_xpath("//div[@aria-controls='rc-tabs-0-panel-PAID']")
    paid_tab.click()
    time.sleep(3)

    # Inserting pending data to sqlite
    next_page_exists = True
    with closing(setup_cursor()) as cur:
        while next_page_exists:     # Waits until table is fully loaded
            try:
                print("LOADING TABLE. . .")
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((
                                                    By.XPATH, "//td[@class='ant-table-cell ant-table-cell-fix-right ant-table-cell-fix-right-first']")))
                table_rows = driver.find_elements_by_xpath("//tbody[@class='ant-table-tbody']/tr")
                print(f"LENGTH OF TABLE: {len(table_rows)}")

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

            except (StaleElementReferenceException, ElementClickInterceptedException) as e:
                print(f"TABLE NOT FULLY LOADED: {e}")
                time.sleep(1)
                pass
    if headless and driver is None:
        print("Quitting headless driver ")
        driver.quit()


def go_order(driver, order_number):
    with closing(setup_cursor()) as cur:
        cur.execute(f"SELECT ginee_order_id FROM orders WHERE order_number = '{order_number}';")
        ginee_order_id = cur.fetchone()[0]
    # Goes to order page
    driver.get(f"https://seller.ginee.com/main/order/order-detail?orderId={ginee_order_id}")


def switch_to_iframe(driver):
    """Switches to iframe to interact with elements"""
    try:
        iframe = WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.ID, 'myIframe')))
        driver.switch_to.frame(iframe)
        print("\tSwitched to iframe")
    except:
        pass


def close_popupwindow(driver):
    close_button = driver.find_elements_by_xpath("//button[@aria-label='Close']")
    order_page_url = 'order/order-detail?orderId=' in driver.current_url

    if close_button and order_page_url:
        try:
            print("\tClosing popup window...")
            close_button[0].click()
            time.sleep(0.25)
        except Exception as e:
            print("\tFailed to close window\n")
            print(e)


def close_tabs(driver, main_window):
    # Closes other tabs & switches to main window
    if len(driver.window_handles) > 1:
        try:
            print("\tClosing other tabs...")
            for window in driver.window_handles:
                if window != main_window:
                    driver.switch_to.window(window)
                    driver.close()
            print("\tSwitching to main tab")
            driver.switch_to.window(main_window)
        except Exception as e:
            print("\tFailed to close tabs\n")
            print(e)


def arrange_shipment(driver):
    print("ARRANGING SHIPMENT")
    switch_to_iframe(driver)
    close_popupwindow(driver)
    try:
        # Arrange Shipment (ready to ship)
        WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//button[normalize-space()='Arrange Shipment']"))).click()
        time.sleep(2)
        WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//button[normalize-space()='Arrange Shipment']"))).click()
        print("\tSuccess.")
    except Exception as e:
        print("\tFailed.\n")
        print(e)


def print_pdf(driver):
    print("PRINTING PDF")
    switch_to_iframe(driver)
    main_window = driver.current_window_handle
    close_tabs(driver, main_window)
    close_popupwindow(driver)
    try:
        print("\tFinding print button")
        # Clicking Print button from the order page
        print_button = driver.find_elements_by_xpath("//button[normalize-space()='Print']")
        if print_button:  # Only available in the order page
            print_button[0].click()
            time.sleep(0.25)

        # Generate AWB pdf in other tab
        print("\tFinding print label button")
        WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                (By.XPATH, "//*[normalize-space()='Print Label']")))
        print_labels = driver.find_elements_by_xpath("//*[normalize-space()='Print Label']")

        for print_label in print_labels:
            try:
                print('Clicking print label')
                print_label.click()
            except ElementClickInterceptedException as e:
                print(e)
                pass

        time.sleep(2)
        order_page_url = 'order/order-detail?orderId=' in driver.current_url

        if order_page_url:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.XPATH, "//td/button[normalize-space()='Print']"))).click()
        else:
            switch_to_iframe(driver)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                                    (By.XPATH, "//button[normalize-space()='Print']"))).click()

        # Switches tab then prints
        print('\tswitching to pdf tab')
        while len(driver.window_handles) == 1:
            time.sleep(2)
        driver.switch_to.window(driver.window_handles[1])           # switches tab
        time.sleep(2)
        driver.execute_script('window.print();')                    # Ctrl + P
        print('\tswitching to print preview')
        while len(driver.window_handles) == 2:
            time.sleep(1.5)
        driver.switch_to.window(driver.window_handles[2])           # switches to print preview
        WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                                        (By.ID, 'selecttrigger-1'))).click()
        driver.find_element_by_xpath(f"//div[@title='{PRINTER}']").click()
        driver.find_element_by_xpath("//button[normalize-space()='Print']").click()
        print("\tPrinted.")
        close_tabs(driver, main_window)
    except Exception as e:
        print("\tFailed.\n")
        print(e)


def arrange_shipment_and_print(driver):
    arrange_shipment(driver)
    print_pdf(driver)


class Application():
    time_out = 3000
    barcode_commands = {    
                        'READY TO SHIP': arrange_shipment,
                        'PRINT': print_pdf,
                        'RTS&P': arrange_shipment_and_print
    }

    def __init__(self, root):
        # Tkinter Configuration
        root.title('Ginee Barcode Automation')
        root.iconbitmap(os.path.join(onedrive_location, 'Ginee', 'ginee-app-logo.ico'))
        root.geometry("400x70")
        self.entry = Entry(root, font=('default', 16))
        self.entry.place(x=10, y=5, width=380, height=50)
        reg = root.register(self.callback)
        self.entry.config(validate="all", validatecommand=(reg, '%P'))
        self.answer = Label(root, text='Please scan barcode', font=(None, 10), bg='white')
        self.answer.pack(pady=30)
        root.after(1000, self.reduce_time_out)
        root.bind_all("<Any-KeyPress>", self.reset_timer)
        root.bind_all("<Any-ButtonPress>", self.reset_timer)
        root.protocol("WM_DELETE_WINDOW", self.close_driver)    # Closes driver on closing of tkinter
        # Initialize Seleniums
        self.driver = self.open_ginee()
        Thread(target=scrape, args=[None, True]).start()

    def open_ginee(self, headless=False):
        print("OPENING GINEE")
        driver = setup_driver(headless=headless, maximized=True, window_position=(-1000, 0))
        login(driver)
        return driver

    def callback(self, input):
        """Verifies if barcode input is valid"""
        input = input.strip()
        print(input)

        if input in self.barcode_commands:
            self.answer.config(text='Command Accepted!')
            order_page_url = self.driver.current_url

            if 'order/order-detail?orderId=' not in order_page_url:
                self.answer.config(text='Please go to order page!', fg='purple')
                return True

            # Execute command scripts
            Thread(target=self.barcode_commands[input], args=[self.driver]).start()
            print('DONE!')
            return True

        # Goes to order page
        elif len(input) >= 12:
            with closing(setup_cursor()) as cur:
                cur.execute(f"SELECT ginee_order_id FROM orders WHERE order_number = '{input}'")
                if cur.fetchone():  # if exists
                    try:
                        go_order(self.driver, input)
                    except WebDriverException:
                        self.answer.config(text=f'PLEASE WAIT: Re-opening Ginee')
                        self.driver.quit()
                        self.driver = self.open_ginee()
                        root.focus_force()      # focuses on window
                        self.entry.focus()
                        go_order(self.driver, input)
                    return True
                else:
                    self.answer.config(text="ORDER NUMBER NOT FOUND", fg='red')
                    return True

        elif input == "":
            self.answer.config(text="Please scan barcode", fg='black')
            return True

        else:
            self.answer.config(text='. . .', fg='blue')
            return True

    def reset_timer(self, event=None):
        # Resets timer to 2 seconds
        if event is not None:
            self.time_out = 2000
        else:
            pass

    def reduce_time_out(self):
        self.time_out = self.time_out-1000
        print(self.time_out)
        root.after(1000, self.reduce_time_out)
        # Clears entry widget every 2 seconds
        if self.time_out == 0:
            print("TIMEOUT REACHES 0")
            self.entry.delete(0, 'end')
            self.reset_timer()
        # Scrapes every 30 minutes of idle
        elif self.time_out % -1800000 == 1:
            print("TIMEOUT REACHES -1800000")
            self.answer.config(text='PLEASE WAIT (SCRAPING)', fg='red')
            Thread(target=scrape, args=(None, True)).start()

    def close_driver(self):
        print("Application closed")
        print("\tClosing driver...")
        self.driver.quit()
        root.destroy()
        print("\tSUCCESS!")


if __name__ == '__main__':
    root = Tk()
    Application(root)
    root.mainloop()
    # scrape(headless=True)
    # driver = setup_driver(headless=True)
    # login(driver)
    # go_order(driver, 420853175910304)
    # print_pdf(driver)
    # driver.quit()
    # scrape(headless=True)
