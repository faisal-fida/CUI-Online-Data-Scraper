import re
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import bs4
import pandas as pd
import math as mt
import captcha_bypass as cb

PATH = "driver/chromedriver.exe"
URL = "https://cuonline.cuiwah.edu.pk:8095/"



def save_html(name, html_page):
    filename = "html/{}.html".format(name)
    print("Saving {}...".format(filename))
    with open(filename, 'w') as file:
        file.write(html_page)


def read_html(name):
    filename = "html/{}.html".format(name)
    with open(filename, "r") as file:
        page = file.read()
    return page


# START
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(PATH, options=chrome_options)
driver.set_page_load_timeout(60)

driver.get(URL)

username = driver.find_element_by_id("MaskedRegNo")
username.send_keys(USERNAME)

password = driver.find_element_by_id("Password")
password.send_keys(PASSWORD)

login_btn = driver.find_element_by_id("LoginSubmit")

while True:
    iframes = driver.find_elements_by_tag_name("iframe")
    for iframe in iframes:
        if iframe.get_attribute("src").startswith("https://www.google.com/recaptcha/api2/anchor"):
            captcha = iframe
    cb.solve_captcha(driver, captcha)

    login_btn.click()
    WebDriverWait(driver, timeout=1000, poll_frequency=1).until(
        EC.staleness_of(login_btn))

    if driver.current_url == "https://cuonline.cuiwah.edu.pk:8095/COURSES":
        print("Login successful!")
        break
    else:
        print("Failed to login!")

# Saving html pages
save_html("courses", driver.page_source)

result_card_btn = driver.find_element_by_id("Result_Card").click()
save_html("result", driver.page_source)
time.sleep(1)

my_profile_btn = driver.find_element_by_id("My_Profile").click()
save_html("profile", driver.page_source)
time.sleep(1)

image = driver.find_element_by_xpath('//div[@id="divImageHolder"]/img')

print("Saving html/picture.png...")
with open('html/picture.png', 'wb') as file:
    file.write(image.screenshot_as_png)

print("Closing browser...")
driver.close()

# Parsing data from HTML pages

# Initializing ExcelWriter
writer = pd.ExcelWriter("data.xlsx", engine='xlsxwriter')
workbook = writer.book

# Profile
page = read_html("profile")
soup = bs4.BeautifulSoup(page, 'html.parser')
data = soup.find("div", attrs={"style": "float:left"})
i = data.find_all("div")
details = {}
for l in i:
    text = l.text.strip()
    name_value = text.split(":")
    details[name_value[0] + ":"] = name_value[1]


details_df = pd.DataFrame(list(details.items()))

table = soup.find(
    "table", class_="table table-striped table-bordered table-hover")
table_df = pd.read_html(str(table), flavor='bs4', converters={
                        x: str for x in range(2)})[0]

# Writing each dataframe from table_container to excel
details_df.to_excel(writer, sheet_name="Profile",
                    index=False, startrow=0, header=None)
table_df.to_excel(writer, sheet_name="Profile", index=False, startrow=8)
worksheet = writer.sheets["Profile"]
cell_format = workbook.add_format()
cell_format.set_bold(True)
worksheet.set_column("A:A", None, cell_format)
worksheet.insert_image("D1", "html/picture.png")
writer.sheets["Profile"].set_column(0, 0, 30)
writer.sheets["Profile"].set_column(1, 1, 35)
writer.sheets["Profile"].set_column(2, 2, 20)

# Courses
page = read_html("courses")
soup = bs4.BeautifulSoup(page, 'html.parser')
table = soup.find(
    "table", class_="table table-striped table-bordered table-hover")
data = pd.read_html(str(table), header=0, flavor='bs4')[0]
data = data.rename(columns={"Attendance Summary": "Attendance"})
for idx in data.index:
    attendance = str(data.loc[idx, "Attendance"])
    percentage = re.search("_percentage: (.*?),", attendance)
    if percentage is None:
        continue
    data.loc[idx, "Attendance"] = int(percentage.group(1))

data.to_excel(writer, sheet_name='Courses', index=False)
for column in data:
    column_width = max(data[column].astype(
        str).map(len).max(), len(column)) + 1
    if mt.isnan(column_width):
        column_width = len(column) + 1
    col_idx = data.columns.get_loc(column)
    writer.sheets["Courses"].set_column(col_idx, col_idx, column_width)

# Result
page = read_html("result")
soup = bs4.BeautifulSoup(page, 'html.parser')
table_container = soup.find_all("div", class_="single_result_container")
# Iterating each table_container which contains further 3 tables for details, marks, cgpa
for idx, container in enumerate(table_container):
    dataframe_list = []
    sheet_name = "Semester {}".format(idx + 1)
    tables = container.find_all("div", class_="table_container")
    # Getting tables from each table container and reading in dataframe saving to dataframe_list
    for table in tables:
        if table.find("table", class_="tbl_one"):
            data = pd.read_html(str(table), header=None, flavor='bs4')[0]
        else:
            data = pd.read_html(str(table), header=1, flavor='bs4')[0]
        dataframe_list.append(data)
    # Writing each dataframe from table_container to excel
    row_num = 0
    for n, dataframe in enumerate(dataframe_list):
        if n == 1:
            dataframe.to_excel(writer, sheet_name=sheet_name,
                               index=False, startrow=row_num)
        else:
            dataframe.to_excel(writer, sheet_name=sheet_name,
                               index=False, startrow=row_num, header=None)
        row_num += dataframe.shape[0] + 2
    writer.sheets[sheet_name].set_column(0, 0, 50)
    writer.sheets[sheet_name].set_column(0, 1, 40)

writer.close()
print("Done!")
