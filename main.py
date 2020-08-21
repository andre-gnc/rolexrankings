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
for i in range(LOOP):
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

# dict_data = {}
col_headers = ('Player ID', 'Player Name', 'Last name', 'Other Name', 'Gender', 'Year', 'Events Played', '1st',
               '2nd', '3rd', '4th-10th', 'MC', 'Year End Rank')

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
irow = 0
icol = 0

for n in range(len(col_headers)):
    worksheet.write(irow, icol, col_headers[n], bold)
    icol += 1

# For each record, find and save the data.
for j in range(len(urls)):
    # Souping each url.
    response = requests.get(urls[j])
    data = response.text
    soup = BeautifulSoup(data, 'html.parser')

    # Saving player's identity.
    print(j + 1)
    slash_position = urls[j].rfind('/')
    player_id = urls[j][slash_position + 1:]
    name = soup.find('small').find_previous().text
    space_position = name.rfind(' ')
    lastname = name[space_position + 1:]
    othername = name[:space_position]

    vcols = [player_id, name, lastname, othername, GENDER]

    # Saving the table.
    tables = soup.find_all('table')
    k = 0
    for table in tables:
        if k != 2:
            k += 1
            continue
        else:
            # Sorting each row out.
            rows = table.find_all('tr')
            m = 0
            for row in rows:
                # Skip the last row (summary).
                if m + 1 == len(rows):
                    break
                # Skip the first row (header).
                elif m == 0:
                    m += 1
                    continue
                else:
                    m += 1
                    irow += 1
                    icol = 0

                    # Write to data.xlsx.
                    for vcol in vcols:
                        worksheet.write(irow, icol, vcol)
                        icol += 1

                    # Sort each column out and save to vcols.
                    columns = row.find_all('td')
                    for column in columns:
                        worksheet.write(irow, icol, column.text)
                        icol += 1
    break
workbook.close()
