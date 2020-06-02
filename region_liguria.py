from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import locale
import re
import pandas as pd


def get_data_liguria(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%d %B')
    date = date.upper()

    # Hide the chrome window
    options = Options()
    options.headless = True
    driver = webdriver.Chrome('C:/webdriver/chromedriver', options=options)

    # Reference website
    driver.get("https://www.sanremonews.it/")

    web_date = driver.find_element_by_xpath("""//*[@id="c6506"]/div/table/tbody/tr/td[11]/h1""")

    death = driver.find_element_by_xpath('//*[@id="c6506"]/div/table/tbody/tr/td[11]/h2')

    prov = [("Imperia", int(death.text))]

    liguria = pd.DataFrame(prov, columns=['provincia', 'decessi'])
    liguria = liguria.sort_values(by=["provincia"])

    liguria
