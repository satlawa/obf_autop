import numpy as np
import pandas as pd

# TODO
# read in data
# filter data
# output it with python docx

class OBFFlaechentabelle(object):

    def __init__(self, data):

        self.data = data

        #---------------------
        #   filter data
        #---------------------

        self.data_wo_fl = self.data[self.data['Schichtalter'] == 0]

        self.data_wo_fl = self.data_wo_fl[['Forstrevier','Abteilung', 'Unterabteil.', \
        'Teilfl.', 'WE-Typ', 'Betriebsklasse', 'Umtriebszeit', 'Nebengrund Art', \
        'Ertragssituation', 'Bewirtschaftungsform', 'Schutzwaldkategorie', \
        'Fläche in HA']]

        # get Forstreviere
        self.fr = np.sort(self.data_wo_fl['Forstrevier'].unique())

        # get Abteilungen for every Forstrevier
        self.fr_abt = []
        for f in self.fr:
            self.fr_abt.append(np.sort(self.data_wo_fl.loc[self.data_wo_fl['Forstrevier']==f,'Abteilung'].unique()))


    def create_table(self, fr, abt):

        # filter one Abteilung
        table = self.data_wo_fl.loc[(self.data_wo_fl['Forstrevier']==fr) & (self.data_wo_fl['Abteilung']==abt)]
        # sort table by WE-Typ , unterabteil. , Teilfl.
        table = table.sort_values(['Abteilung', 'WE-Typ', 'Unterabteil.', 'Teilfl.'], ascending=[True, False, True, True])

        table['fl_ww'] = table.loc[(table['Bewirtschaftungsform']=='W'),'Fläche in HA']
        table['fl_sw'] = table.loc[(table['Bewirtschaftungsform']=='S'),'Fläche in HA']
        table['fl_nhb'] = table.loc[(table['Nebengrund Art']==3) | (table['Nebengrund Art']==4),'Fläche in HA']
        table['fl_ngp'] = table.loc[(table['Nebengrund Art']==5) | (table['Nebengrund Art']==6),'Fläche in HA']
        table['fl_ngu'] = table.loc[(table['Nebengrund Art']<3) | (table['Nebengrund Art']>6),'Fläche in HA']

        sum_fl = {'Fläche in HA': table['Fläche in HA'].sum(), 'fl_ww': table['fl_ww'].sum(), 'fl_sw': table['fl_sw'].sum(), 'fl_nhb': table['fl_nhb'].sum(), 'fl_ngp': table['fl_ngp'].sum(), 'fl_ngu': table['fl_ngu'].sum()}
        sum_df = pd.DataFrame(data=sum_fl, index = ["Ges."])
        table = table.append(sum_df, sort=False)
        table.columns = ['FR', 'Abt', 'UAbt', 'Tfl', 'WE Typ', 'BKL', 'UZ', 'NG Typ', 'EtrS', \
        'BW Typ', 'SW Typ', 'Fläche Ges', 'Fläche WW', 'Fläche SW', 'Fläche NHB', 'Fläche pNG', 'Fläche uNG']
        table = table.reindex(columns= ['Abt', 'UAbt', 'Tfl', 'WE Typ', 'BKL', 'UZ', 'EtrS', \
        'BW Typ', 'SW Typ', 'NG Typ', 'Fläche WW', 'Fläche SW', 'Fläche NHB', 'Fläche pNG', 'Fläche uNG', 'Fläche Ges'])

        table = table.round(2)

        #table["FR"] = pd.to_numeric(table["FR"])
        table["Abt"] = table["Abt"].astype(int)
        table["Tfl"] = table["Tfl"].astype(int)
        table["BKL"] = table["BKL"].astype(int)
        table["UZ"] = table["UZ"].astype(int)
        table["NG Typ"] = table["NG Typ"].astype(int)

        #table["Abt"] = pd.to_numeric(table["Abt"])
        #table["Tfl"] = pd.to_numeric(table["Tfl"])
        #table["BKL"] = pd.to_numeric(table["BKL"])
        #table["UZ"] = pd.to_numeric(table["UZ"])
        #table["NG Typ"] = pd.to_numeric(table["NG Typ"])
        #table = table.fillna('-')

        return table
