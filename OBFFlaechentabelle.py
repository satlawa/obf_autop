import numpy as np
import pandas as pd

from OBFDictionary import OBFDictionary

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

        self.dic = OBFDictionary()
        self.dictionary()

        # get Abteilungen for every Forstrevier
        self.fr_abt = []
        for f in self.fr:
            self.fr_abt.append(np.sort(self.data_wo_fl.loc[self.data_wo_fl['Forstrevier']==f,'Abteilung'].unique()))


    def dictionary(self):

        # get Teiloperat
        self.dic.set_to(self.data['Teiloperats-ID'].unique()[0])
        # get Forstbetrieb
        self.dic.set_fb(self.data['Forstbetrieb'].unique()[0])

        # get Laufzeit
        #x = self.data['Beg. Laufzeit'].iloc[0].to_pydatetime()
        x = self.data['Beg. Laufzeit'].iloc[0][-4:]
        y = self.data['Ende Laufzeit'].iloc[0][-4:]
        self.dic.set_laufzeit_begin(x)
        self.dic.set_laufzeit_end(y)
        self.dic.set_laufzeit(x + '-' + y)

        x = np.sort(self.data['Forstrevier'].unique())
        self.dic.set_fr(x[x.nonzero()])
        x = np.sort(self.data['Umtriebszeit'].unique())
        self.dic.set_uz(x[x.nonzero()])
        x = np.sort(self.data['Nutzdringlichkeit'].unique())
        self.dic.set_dring(x[x.nonzero()])
        x[x.nonzero()]

        # set fr dict
        self.dic.set_dic_fr()


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

        sum_fl = {'Abteilung': 'Ges.', \
            'Fläche in HA': table['Fläche in HA'].sum(), 'fl_ww': table['fl_ww'].sum(), \
            'fl_sw': table['fl_sw'].sum(), 'fl_nhb': table['fl_nhb'].sum(), \
            'fl_ngp': table['fl_ngp'].sum(), 'fl_ngu': table['fl_ngu'].sum()}

        table = table.fillna('-')
        table[['Abteilung', 'Unterabteil.', 'Teilfl.', 'WE-Typ', 'Betriebsklasse', 'Umtriebszeit', 'Nebengrund Art']] = \
        table[['Abteilung', 'Unterabteil.', 'Teilfl.', 'WE-Typ', 'Betriebsklasse', 'Umtriebszeit', 'Nebengrund Art']].astype(str)

        table['Betriebsklasse'] = table['Betriebsklasse'].str.replace('.0', '', regex=False)
        table['Nebengrund Art'] = table['Nebengrund Art'].str.replace('.0', '', regex=False)

        sum_df = pd.DataFrame(data=sum_fl, index = ["Ges."])
        sum_df = sum_df.round(2)
        table = table.append(sum_df, sort=False)
        table.columns = ['FR', 'Abt', 'UAbt', 'TF', 'WE', 'BKL', 'UZ', 'NG', 'ES', \
        'BW', 'SW', 'Fl Ges', 'Fl WW', 'Fl SW', 'Fl NHB', 'Fl pNG', 'Fl uNG']
        table = table.reindex(columns= ['Abt', 'UAbt', 'TF', 'WE', 'BKL', 'UZ', 'ES', \
        'BW', 'SW', 'NG', 'Fl WW', 'Fl SW', 'Fl NHB', 'Fl pNG', 'Fl uNG', 'Fl Ges'])

        #table = table.round(2)

        table = table.fillna('-')

        return table
