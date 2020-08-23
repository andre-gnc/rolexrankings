import os
import time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import requests
from bs4 import BeautifulSoup
import xlsxwriter


TIMEOUT = 20
LOOP = 1
GENDER = 'Female'


def load_records(lr_loop, lr_wait, lr_window_height_half):
    # Loading all records of players.
    for index_button in range(lr_loop):
        try:
            button = lr_wait.until(ec.element_to_be_clickable((By.XPATH, '/html/body/app/main/div[2]/div[4]/button')))

            element_coordinate_y = button.location['y']
            driver.execute_script("window.scrollTo(0, %d)" % (element_coordinate_y - lr_window_height_half))
            # time.sleep(1)
            button.click()
            time.sleep(1)
        except TimeoutException:
            break


def make_soup(ms_url):
    # Make soups of each url.
    response = requests.get(ms_url)
    data = response.text
    ms_soup = BeautifulSoup(data, 'html.parser')
    return ms_soup


def selenium_gets_urls(sgu_driver):
    # Finding and saving urls.
    sgu_elements_url = sgu_driver.find_elements_by_css_selector("td._2Dx28 [href]")
    sgu_urls = []
    for sgu_element_url in sgu_elements_url:
        urls.append(sgu_element_url.get_attribute('href'))
    return sgu_urls


def xlsxwriter_write(xw_items, xw_index_row, xw_worksheet, xw_start, xw_format):
    for xw_counter_col, xw_item in enumerate(xw_items, xw_start):
        xw_worksheet.write(xw_index_row, xw_counter_col, xw_item, xw_format)


def make_driver_chrome():
    mdc_driver = webdriver.Chrome(executable_path=os.path.abspath('chromedriver.exe'))
    mdc_driver.get('https://www.rolexrankings.com/rankings/2020-08-10')
    return mdc_driver


driver = make_driver_chrome()

wait = WebDriverWait(driver, TIMEOUT)

window_size = driver.get_window_size()
window_height = window_size['height']
window_height_half = int(window_height / 2)

# Loading all records of players.
load_records(LOOP, wait, window_height_half)

# Finding and saving urls.
urls = []
url = selenium_gets_urls(driver)

col_headers = ('Player ID', 'Player Name', 'Last name', 'Other Name', 'Gender', 'Year', 'Events Played', '1st',
               '2nd', '3rd', '4th-10th', 'MC', 'Year End Rank')

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
index_row = 0

# Write the headers.
xlsxwriter_write(col_headers, index_row, worksheet, 0, bold)

# For each record, find and save the data.
for index_urls, url in enumerate(urls):
    # Make soups of each url.
    soup = make_soup(url)

    # Monitor its progress.
    print(index_urls + 1)

    # Saving player's identity.
    slash_position = url.rfind('/')
    player_id = url[slash_position + 1:]
    name = soup.find('small').find_previous().text
    position_space = name.rfind(' ')
    lastname = name[position_space + 1:]
    othername = name[:position_space]

    value_cols = [player_id, name, lastname, othername, GENDER]

    # Saving the table.
    tables = soup.find_all('table')
    for index_table, table in enumerate(tables, 1):
        # The required table is the 3rd.
        if index_table != 3:
            continue

        rows = table.find_all('tr')

        # Remove the last row (Tot).
        rows.pop()

        for counter_row, row in enumerate(rows):
            # The required rows are from the 2nd.
            # Skip the 0 indexed.
            if counter_row == 0:
                continue

            index_row += 1

            # Write the values of the first 5 columns.
            xlsxwriter_write(value_cols, index_row, worksheet, 0, None)

            column_urls = row.find_all('td')
            columns = []
            # Get the content as text from each column.
            for column_url in column_urls:
                column_text = column_url.text
                columns.append(column_text)

            # Write the last values of other columns from the 6th.
            # It's 5 since start at 0.
            xlsxwriter_write(columns, index_row, worksheet, 5, None)
    break
workbook.close()
