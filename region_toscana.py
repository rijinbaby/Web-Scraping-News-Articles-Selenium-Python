from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import locale
import re
import pandas as pd


def get_data_toscana(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%d %B')

    # Hide the chrome window
    options = Options()
    options.headless = True
    driver = webdriver.Chrome('C:/webdriver/chromedriver', options=options)

    # Reference website
    driver.get("https://www.toscana-notizie.it/home")

    # Title contents
    article_title = ["Coronavirus", "nuovi casi", "decessi", "guarigioni"]

    posts = driver.find_elements_by_xpath(
        """//*[@id="portlet_com_liferay_asset_publisher_web_portlet_AssetPublisherPortlet_INSTANCE_bfnpGlY0o8Sy"]/div/div/div/div[2]/div/section/div/div/div[2]/a""")

    # get position of link of matching title
    pos = 0
    for x in posts:
        pos += 1
        a_date = driver.find_element_by_xpath(
            '//*[@id="portlet_com_liferay_asset_publisher_web_portlet_AssetPublisherPortlet_INSTANCE_bfnpGlY0o8Sy"]/div/div/div/div[2]/div/section/div[' + str(
                pos) + ']/div/div[1]/div[1]')
        result = (all(p in x.text for p in article_title)) and (date in a_date.text)
        if result is True:
            break

    pos = str(pos)

    # getting the link of article
    href = driver.find_element_by_xpath(
        '//*[@id="portlet_com_liferay_asset_publisher_web_portlet_AssetPublisherPortlet_INSTANCE_bfnpGlY0o8Sy"]/div/div/div/div[2]/div/section/div[' + pos + ']/div/div[2]/a')
    link = href.get_attribute('href')

    # content of the article needed
    driver.get(link)
    news = driver.find_element_by_class_name("rt-page__content-text")
    content = news.text

    # positioning death info
    res = re.findall(r'([^.]*Relativamente alla provincia di notifica[^.]*)', content)[0]
    res = res.split(',')

    prov = []
    for x in res:
        if "Arezzo" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Arezzo", int(death)])

        if "Firenze" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Firenze", int(death)])

        if "Grosseto" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Grosseto", int(death)])

        if "Livorno" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Livorno", int(death)])

        if "Lucca" in x:
            loc = re.search("\d", x).start()
            death = x[loc:loc + 1]
            prov.append(["Lucca", int(death)])

        if "Massa Carrara" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Massa Carrara", int(death)])

        if "Pisa" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Pisa", int(death)])

        if "Pistoia" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Pistoia", int(death)])

        if "Prato" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Prato", int(death)])

        if "Siena" in x:
            loc = re.search("\d", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Siena", int(death)])

    toscana_prov = ["Arezzo", "Firenze", "Grosseto", "Livorno", "Lucca", "Massa Carrara", "Pisa", "Pistoia", "Prato",
                    "Siena"]

    # check if all province present
    for x in toscana_prov:
        if x in str(prov):
            continue
        else:
            prov.append([x, int(0)])

    toscana = pd.DataFrame(prov, columns=['provincia', 'decessi'])
    toscana = toscana.sort_values(by=["provincia"])

    driver.close()
    return(toscana,link)
