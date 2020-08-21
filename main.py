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

driver = webdriver.Chrome(executable_path=os.path.abspath('chromedriver.exe'))
driver.get('https://www.rolexrankings.com/rankings/2020-08-10')

wait = WebDriverWait(driver, TIMEOUT)

window_size = driver.get_window_size()
window_height = window_size['height']
window_height_half = int(window_height / 2)

# Loading all records of players.
for index_button in range(LOOP):
    try:
        button = wait.until(ec.element_to_be_clickable((By.XPATH, '/html/body/app/main/div[2]/div[4]/button')))

        element_coordinate_y = button.location['y']
        driver.execute_script("window.scrollTo(0, %d)" % (element_coordinate_y - window_height_half))
        # time.sleep(1)
        button.click()
        time.sleep(1)
    except TimeoutException:
        break

# Finding and saving urls.
elements_url = driver.find_elements_by_css_selector("td._2Dx28 [href]")
urls = []
for element_url in elements_url:
    urls.append(element_url.get_attribute('href'))

col_headers = ('Player ID', 'Player Name', 'Last name', 'Other Name', 'Gender', 'Year', 'Events Played', '1st',
               '2nd', '3rd', '4th-10th', 'MC', 'Year End Rank')

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
index_row = 0
# index_col = 0

# Write the headers.
for counter_col, col_header in enumerate(col_headers):
    worksheet.write(index_row, counter_col, col_header, bold)

# For each record, find and save the data.
for index_urls, url in enumerate(urls):
    # Souping each url.
    response = requests.get(url)
    data = response.text
    soup = BeautifulSoup(data, 'html.parser')

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
            for counter_col, value_col in enumerate(value_cols):
                worksheet.write(index_row, counter_col, value_col)

            columns = row.find_all('td')
            # Write the last values of other columns from the 6th.
            # It's 5 since start at 0.
            for counter_col, column in enumerate(columns, 5):
                worksheet.write(index_row, counter_col, column.text)
    break
workbook.close()
