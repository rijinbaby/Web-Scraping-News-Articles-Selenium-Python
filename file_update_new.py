import glob, os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import locale
from pycel import ExcelCompiler

class Update_Excel():
    """
    Updates the excel files for the provinces in
    a specified region with the specified data.

    Parameters
    ----------
    path : string, default=None
        The location of the folder containing
        the excel files.

    regione : string, {""}
        The name of the region to update
        
    csv_pcm : string
        The path of the csv from the Protezione Civile. This overrides the automatic download
    """
    def __init__(self, region, path=None):
        
        self.path = path
        self.region = region
        self.region_list = ["Abruzzo", "Basilicata", "Calabria", "Emilia-Romagna", "Liguria", "Molise", "Sicilia",
                            "Toscana", "Umbria"]
        locale.setlocale(locale.LC_TIME, 'it_IT.UTF-8')
        
        #population data
        self.pop_prov = {
            "Agrigento": 434870, "Arezzo": 342654, "Bologna": 1014619, "Caltanissetta": 262458, 
             "Campobasso": 221238, "Catania": 1107702,"Catanzaro": 358316, "Chieti": 385588, 
             "Cosenza": 705753, "Crotone": 174980, "Enna": 164788, "Ferrara": 345691, "Firenze": 1011349, 
             "Forlì-Cesena": 394627, "Grosseto": 221629, "Imperia": 213840, "Isernia": 84379,"L'Aquila":299031,
             "Livorno": 334832, "Lucca": 387876, "Massa Carrara": 194878, "Matera": 197909, 
             "Messina": 626876, "Modena": 705393, "Palermo": 1252588, "Parma": 451631, 
             "Perugia": 656382, "Pescara": 318909, "Piacenza": 287152, "Pisa": 419037, 
             "Pistoia": 292473, "Potenza": 364960, "Prato": 257716, "Ragusa": 320893, "Ravenna": 389456, 
             "Reggio di Calabria": 548009,"Reggio nell'Emilia":531891, "Rimini": 339017, "Savona": 276064, "Siena": 267197, 
             "Siracusa": 399224, "Teramo": 308052, "Terni": 225633, "Trapani": 430492, 
             "Vibo Valentia": 160073   
        }
        
        self.csv_columns = ["Province","Day","Date","New_cases","Curr_pos_cases","Deaths",
                            "Tot_deaths","Lethality_rate","Cumul_rates"]
        
    def _get_default_path(self):
        """
        Get the corresponding default path for the
        excel files of the specified region.
        """

#         if self.use_csv:
        if self.region == "Abruzzo":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Abruzzo"
        elif self.region == "Basilicata":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Basilicata"
        elif self.region == "Calabria":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Calabria"
        elif self.region == "Emilia-Romagna":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Emilia-Romagna"
        elif self.region == "Liguria":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Liguria"
        elif self.region == "Molise":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Molise"
        elif self.region == "Sicilia":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Sicilia"
        elif self.region == "Toscana":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Toscana"
        elif self.region == "Umbria":
            return "C:/Users/Rijin/Documents/GitHub/COVID-Pro-Dataset/Deaths_Umbria"
        else:
            raise ValueError(self.region + " is not valid. " \
                                           "Please choose among " + ' '.join([str(elem) for elem in self.region_list]))

    def _get_lista_prov(self):
        """
        Get the corresponding list of provinces
        from the specified region.
        """
        if self.region == "Abruzzo":
            return ["Chieti", "L'Aquila", "Pescara", "Teramo"]
        elif self.region == "Basilicata":
            return ["Matera", "Potenza"]
        elif self.region == "Calabria":
            return ["Catanzaro", "Cosenza", "Crotone", "Reggio di Calabria", "Vibo Valentia"]
        elif self.region == "Emilia-Romagna":
            return ["Bologna", "Ferrara", "Forlì-Cesena", "Modena", "Parma", "Piacenza", "Ravenna",
                    "Reggio nell'Emilia", "Rimini"]
        elif self.region == "Liguria":
            return ["Imperia","Savona"]
        elif self.region == "Molise":
            return ["Campobasso", "Isernia"]
        elif self.region == "Sicilia":
            return ["Agrigento", "Caltanissetta", "Catania", "Enna", "Messina", "Palermo", "Ragusa", "Siracusa",
                    "Trapani"]
        elif self.region == "Toscana":
            return ["Arezzo", "Firenze", "Grosseto", "Livorno", "Lucca", "Massa Carrara", "Pisa", "Pistoia", "Prato",
                    "Siena"]
        elif self.region == "Umbria":
            return ["Perugia", "Terni"]
        else:
            raise ValueError(self.region + " is not valid. " \
                                           "Please choose among " + ' '.join([str(elem) for elem in self.region_list]))

    def _get_files_list(self):
        """
        Enumerates the files in the folder and retrieves
        their full paths.

        Returns
        -------
        files : list
            The list of full paths for the files
            in the folder.
        """

        if self.path is None:
            self.path = self._get_default_path()

        if not str(self.path).endswith("/"):
            self.path = str(self.path) + "/"

        files = glob.glob(self.path + '*.csv')
        files.sort()

        return files

    def _get_decessi_casi(self, df, i, prov, csv_prov_github):
        """
        Get the number of deaths and infected

        Parameters
        ----------
        df : Pandas DataFrame
            The DataFrame containing the new data.

        i : int
            Counter

        prov : string
            The name of the province

        csv_prov_github : Pandas DataFrame
            The DataFrame containing data from the Protezione
            Civile.

        Returns
        -------
        tot_casi : int
            The cumulative number of infected.

        decessi : int
            The number of deaths. It can either be
            the daily or cumulative count.
        """

        tot_casi, decessi = None, None

        if self.region in ["Calabria", "Sicilia"]:
            # Casi_att_positivi
            tot_casi = csv_prov_github[csv_prov_github["denominazione_provincia"] == prov]["totale_casi"].values[0]
            # decessi_tot
            decessi_tot = df.iloc[i]["decessi_tot"]
            return tot_casi, decessi_tot

        else:
            tot_casi = csv_prov_github[csv_prov_github["denominazione_provincia"] == prov]["totale_casi"].values[0]
            decessi = df.iloc[i]["decessi"]
            return tot_casi, decessi
        

    def update_csv(self, df, day_spec):
        """
        Updates the excel files with the specified data for all the provinces
        of the region

        Parameters
        ----------
        df : Pandas DataFrame
            The DataFrame containing the new data.
        """

        self.region = str(self.region)

        if self.region not in self.region_list:
            raise ValueError(self.region + " is not valid. " \
                                            "Please choose among " + ' '.join([str(elem) for elem in self.region_list]))

        self.prov_list = self._get_lista_prov()

        files = self._get_files_list()

        df = df.sort_values(by=["provincia"])

        date = datetime.strftime(datetime.now() - timedelta(day_spec), "%Y%m%d")

        # Scarica il csv province della PCM
        csv_prov_github = pd.read_csv(
            "https://raw.githubusercontent.com/pcm-dpc/COVID-19/master/dati-province/dpc-covid19-ita-province.csv",
            error_bad_lines=False)
        csv_prov_github['data'] = pd.to_datetime(csv_prov_github.data)
        csv_prov_github['data'] = csv_prov_github['data'].dt.strftime('%Y%m%d')
        csv_prov_github = csv_prov_github[(csv_prov_github['data'] == date)]
        
        
        if self.region in ["Calabria", "Sicilia"]:
            for i, file in enumerate(files):
                prov = self.prov_list[i]
                tot_casi, decessi_tot = self._get_decessi_casi(df, i, prov, csv_prov_github)
                self._write_csv_CnS(file, decessi_tot, tot_casi,prov)
                print(prov + ": done")

        else:
            for i, file in enumerate(files):
                prov = self.prov_list[i]
                tot_casi, decessi = self._get_decessi_casi(df, i, prov, csv_prov_github)
                self._write_csv(file, decessi, tot_casi,prov)
                print(prov + ": done")
        
        
    def _write_csv(self, file, decessi, tot_casi,prov):
        """
        Updates the excel file for the specified province

        Parameters
        ----------
        file : string
            The file path.

        decessi : int
            The number of deaths. It can either be
            the daily or cumulative count.

        tot_casi : int
            The cumulative number of infected.
        """
        # Read data
        csv_df = pd.read_csv(file)
        last_line = csv_df.iloc[-1]

        # Province
        prov1 = last_line["Province"]

        # Day
        day = int(last_line["Day"]) + 1

        # Data
        data = datetime.strptime(last_line["Date"], '%Y-%m-%d') + timedelta(1)
        data = datetime.strftime(data, '%Y-%m-%d')

        # Current positive cases
        cpc = tot_casi

        # New cases
        nc = cpc - int(last_line["Curr_pos_cases"])

        # Deaths and total deaths
        deaths = decessi
        tot_deaths = deaths + int(last_line["Tot_deaths"])

        # Lethality rate
        let_rate = tot_deaths / cpc

        # Incidence rate
        pop_prov = self.pop_prov[prov]
        incid = tot_casi/pop_prov*100000

        # Append new data
        new_data = pd.DataFrame([[prov1, day, data, nc, cpc, deaths, tot_deaths, let_rate, incid]], columns=csv_df.columns)
        df_final = pd.concat([csv_df, new_data])
        df_final.reset_index(drop=True, inplace=True)

        # Write to file
        df_final.to_csv(file, index=False, encoding="utf-8")
        
        
    def _write_csv_CnS(self, file, decessi_tot, tot_casi,prov):
        """
        Updates the excel file for the specified province

        Parameters
        ----------
        file : string
            The file path.

        decessi : int
            The number of deaths. It can either be
            the daily or cumulative count.

        tot_casi : int
            The cumulative number of infected.
        """
        # Read data
        csv_df = pd.read_csv(file)
        last_line = csv_df.iloc[-1]

        # Province
        prov1 = last_line["Province"]

        # Day
        day = int(last_line["Day"]) + 1

        # Data
        data = datetime.strptime(last_line["Date"], '%Y-%m-%d') + timedelta(1)
        data = datetime.strftime(data, '%Y-%m-%d')

        # Current positive cases
        cpc = tot_casi

        # New cases
        nc = cpc - int(last_line["Curr_pos_cases"])

        # Deaths and total deaths     
        tot_deaths = decessi_tot
        deaths = tot_deaths - int(last_line["Tot_deaths"])

        # Lethality rate
        let_rate = tot_deaths / cpc

        # Incidence rate
        pop_prov = self.pop_prov[prov]
        incid = tot_casi/pop_prov*100000

        # Append new data
        new_data = pd.DataFrame([[prov1, day, data, nc, cpc, deaths, tot_deaths, let_rate, incid]], columns=csv_df.columns)
        df_final = pd.concat([csv_df, new_data])
        df_final.reset_index(drop=True, inplace=True)

        # Write to file
        df_final.to_csv(file, index=False, encoding="utf-8")
         