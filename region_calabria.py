from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import locale
import re
import pandas as pd


def get_data_calabria(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%d/%m/%Y')

    # Hide the chrome window
    options = Options()
    options.headless = True
    driver = webdriver.Chrome('C:/webdriver/chromedriver', options=options)

    # Reference website
    driver.get("https://portale.regione.calabria.it/website/ugsp/ufficiostampa/calabrianotizie/")

    # Title contents
    article_title = "BOLLETTINO DELLA REGIONE CALABRIA DEL " + date

    # Reference website
    driver.get("https://portale.regione.calabria.it/website/ugsp/ufficiostampa/calabrianotizie/")

    posts = driver.find_elements_by_xpath("/html/body/div[3]/div[3]/div[2]/div[3]/div/div/h3/a")

    # get position of link of matching title
    pos = 0
    for x in posts:
        pos += 1
        result = (article_title == x.text)
        if result is True:
            break

    pos = str(pos)

    # getting the link of article
    href = driver.find_element_by_xpath('/html/body/div[3]/div[3]/div[2]/div[3]/div[' + pos + ']/div/h3/a')
    link = href.get_attribute('href')

    # content of the article needed
    driver.get(link)

    news = driver.find_element_by_class_name("fullnewsview")
    content = news.text

    res = re.findall(r'([^.]*deceduti[^.]*)', content)
    res = [x.lower() for x in res]

    prov = []
    for x in res:
        if "catanzaro" in x:
            loc = re.search("\d deceduti", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Catanzaro", int(death)])
        if "cosenza" in x:
            loc = re.search("\d deceduti", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Cosenza", int(death)])
        if "reggio calabria" in x:
            loc = re.search("\d deceduti", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Reggio Calabria", int(death)])
        if "crotone" in x:
            loc = re.search("\d deceduti", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Crotone", int(death)])
        if "vibo valentia" in x:
            loc = re.search("\d deceduti", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Vibo Valentia", int(death)])

    calabria = pd.DataFrame(prov, columns=['provincia', 'decessi_tot'])
    calabria = calabria.sort_values(by=["provincia"])

    driver.close()
    return (calabria, link)
