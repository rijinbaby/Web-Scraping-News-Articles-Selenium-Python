from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
import locale
import re
import pandas as pd


def get_data_emilia_romagna(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%d/%m/%Y')

    # Hide the chrome window
    options = Options()
    options.headless = True
    driver = webdriver.Chrome('C:/webdriver/chromedriver', options=options)

    # Reference website
    driver.get("https://www.regione.emilia-romagna.it/notizie")

    # article_titles
    article_title_1 = ["coronavirus", "aggiornamento", "positivi"]
    article_title_2 = ["coronavirus", "positivi", "guariti"]
    article_title_3 = ["coronavirus", "positivi", "l'aggiornamento"]

    posts = driver.find_elements_by_xpath("""//*[@id="content-core"]/article/h2/a""")
    # for post in posts:
    #     print(post.text)

    # get position of link of matching title
    pos = 0
    for x in posts:
        pos += 1
        a_date = driver.find_element_by_xpath('//*[@id="content-core"]/article[' + str(pos) + ']/div[4]/span')
        result = ((all(p in x.text.lower() for p in article_title_1)) or (
            all(p in x.text.lower() for p in article_title_2))
                  or (all(p in x.text.lower() for p in article_title_3))) and (date in a_date.text)
        if result is True:
            break

    pos = str(pos)

    # getting the link of article
    href = driver.find_element_by_xpath('//*[@id="content-core"]/article[' + pos + ']/h2/a')
    link = href.get_attribute('href')

    # content of the article needed
    driver.get(link)
    news = driver.find_element_by_class_name("news-text")
    content = news.text

    res = re.findall(r'([^.]*nuovi decessi[^.]*)', content)[1]
    res = res.split(',')

    prov = []
    for x in res:
        if "Piacenza" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Piacenza", int(death)])
            else:
                prov.append(["Piacenza", int(0)])

        if "Parma" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Parma", int(death)])
            else:
                prov.append(["Parma", int(0)])

        if "Reggio Emilia" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Reggio nell'Emilia", int(death)])
            else:
                prov.append(["Reggio nell'Emilia", int(0)])

        if "Modena" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Modena", int(death)])
            else:
                prov.append(["Modena", int(0)])

        if "Bologna" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Bologna", int(death)])
            else:
                prov.append(["Bologna", int(0)])

        if "Ferrara" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Ferrara", int(death)])
            else:
                prov.append(["Ferrara", int(0)])

        if "Ravenna" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Ravenna", int(death)])
            else:
                prov.append(["Ravenna", int(0)])

        if "Forlì-Cesena" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Forlì-Cesena", int(death)])
            else:
                prov.append(["Forlì-Cesena", int(0)])

        if "Rimini" in x:
            loc = re.search("\d", x)
            if loc is not None:
                loc1 = re.search("\d", x).start()
                death = x[loc1 - 1:loc1 + 1]
                prov.append(["Rimini", int(death)])
            else:
                prov.append(["Rimini", int(0)])

    emilia_romagna_prov = ["Bologna", "Ferrara", "Forlì-Cesena", "Modena", "Parma", "Piacenza", "Ravenna",
                           "Reggio nell'Emilia", "Rimini"]

    # check if all province present
    for x in emilia_romagna_prov:
        if x in str(prov):
            continue
        else:
            prov.append([x, int(0)])

    emilia_romagna = pd.DataFrame(prov, columns=['provincia', 'decessi'])
    emilia_romagna = emilia_romagna.sort_values(by=["provincia"])

    driver.close()
    return(emilia_romagna,link)