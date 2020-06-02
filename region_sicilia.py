from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import locale
import re
import pandas as pd


def get_data_sicilia(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%d-%b-%Y')
    date = date.upper()

    # Hide the chrome window
    options = Options()
    options.headless = True
    driver = webdriver.Chrome('C:/webdriver/chromedriver', options=options)

    # Reference website
    driver.get("http://pti.regione.sicilia.it/portal/page/portal/PIR_PORTALE/PIR_ArchivioLaRegioneInforma")

    # Reference title
    article_title1 = "Così l'aggiornamento nelle nove province della Sicilia"
    article_title2 = "L'aggiornamento nelle varie province"

    posts = driver.find_elements_by_xpath("""//*[@id="rg421663"]/tbody/tr/td/div/div""")
    # for post in posts:
    #     print(post.text)

    # get position of link of matching title
    pos = 0
    for x in posts:
        pos += 1
        result = (date in x.text) and ((article_title1 in x.text) or (article_title2 in x.text))
        if result is True:
            break

    pos = str(pos)

    # getting the link of article
    href = driver.find_element_by_xpath('//*[@id="rg421663"]/tbody/tr[' + pos + ']/td/div/div/div/div/div[2]/a')
    link = href.get_attribute('href')

    # content of the article needed
    driver.get(link)
    news = driver.find_element_by_class_name("TitoloGr_FotoUpGr_Testo")
    content = news.text

    # extracting death numbers
    res = re.findall(r'([^.]*deceduto[^.]*)', content)[0]
    res = res.split(';')
    res = [x.lower() for x in res]

    prov = []
    for x in res:
        if "agrigento" in x:
            loc = re.search("\d deceduto", x).start()
            death = x[loc - 1:loc + 1]
            prov.append(["Agrigento", int(death)])
        if "caltanissetta" in x:
            death = x[-3:-1]
            prov.append(["Caltanissetta", int(death)])
        if "catania" in x:
            death = x[-4:-1]
            prov.append(["Catania", int(death)])
        if "enna" in x:
            death = x[-3:-1]
            prov.append(["Enna", int(death)])
        if "messina" in x:
            death = x[-3:-1]
            prov.append(["Messina", int(death)])
        if "palermo" in x:
            death = x[-3:-1]
            prov.append(["Palermo", int(death)])
        if "ragusa" in x:
            death = x[-3:-1]
            prov.append(["Ragusa", int(death)])
        if "siracusa" in x:
            death = x[-3:-1]
            prov.append(["Siracusa", int(death)])
        if "trapani" in x:
            death = x[-3:-1]
            prov.append(["Trapani", int(death)])

    siciliana = pd.DataFrame(prov, columns=['provincia', 'decessi_tot'])
    siciliana = siciliana.sort_values(by=["provincia"])
    siciliana_prov = ["Agrigento", "Caltanissetta", "Catania", "Enna", "Messina", "Palermo", "Ragusa", "Siracusa",
                      "Trapani"]

    driver.close()
    return(siciliana,link)