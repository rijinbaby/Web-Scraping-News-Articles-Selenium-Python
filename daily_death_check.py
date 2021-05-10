import pandas as pd
from datetime import datetime, timedelta

def daily_death(day_spec):
    # setting date
    date = datetime.strftime(datetime.now() - timedelta(day_spec), '%Y-%m-%d')
    yesterday = datetime.strftime(datetime.now() - timedelta(day_spec + 1), '%Y-%m-%d')

    #data from github
    csv_reg_github = pd.read_csv(
        "https://raw.githubusercontent.com/pcm-dpc/COVID-19/master/dati-regioni/dpc-covid19-ita-regioni.csv",error_bad_lines=False)
    csv_reg_github['data'] = pd.to_datetime(csv_reg_github.data)
    csv_reg_github['data'] = csv_reg_github['data'].dt.strftime('%Y-%m-%d')

    region_list = ["Abruzzo", "Basilicata", "Calabria", "Emilia-Romagna", "Liguria", "Molise", "Sicilia", "Toscana",
                   "Umbria"]
    death_df = []

    for reg in region_list:
        death = int(csv_reg_github.loc[(csv_reg_github['denominazione_regione'] == reg) & (
                    csv_reg_github['data'] == date)].deceduti) - int(csv_reg_github.loc[(csv_reg_github[
                                                                                             'denominazione_regione'] == reg) & (
                                                                                                    csv_reg_github[
                                                                                                        'data'] == yesterday)].deceduti)
        death_df.append([reg, death])

    death_df = pd.DataFrame(death_df, columns=['Region', 'Death'])

    return(death_df,date)
