#
#   fuc_tbl_naturschutz (typ, merk=0)
#   fuc_tbl_biotop ()
#   fuc_tbl_birdlife ()
#   fuc_tbl_saatgut ()
#

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from OBFDictionary import OBFDictionary

class OBFNatur(object):

    def __init__(self, data):

        # set 0% to 100%
        data.loc[data['Flächenanteil']==0, 'Flächenanteil'] = 100
        # calc the real area = flaeche * flaechenanteil
        data['Fläche in HA'] = data['Fläche in HA'] * data['Flächenanteil']/100

        # save data in class
        self.data = data
        # fill nan with 0
        #self.data = self.data.fillna(0)


    def printit(self):
        print(self.data)
        return(self.data)


    def set_dic(self, path):
        # load data from excel
        dic_auswk = pd.read_excel(path, 'Auswertekategorie')
        # load data from excel
        dic_merk = pd.read_excel(path, 'Merkmal')

        # set dic from excel data
        self.dic_auswkGruppe = dic_auswk.set_index('Abk')['Gruppe'].to_dict()
        self.dic_auswkTyp = dic_auswk.set_index('Abk')['Typ'].to_dict()
        self.dic_merk = dic_merk.set_index('Code')['Typ'].to_dict()

        # add missing data -> add column Gruppe
        self.data['AuswGruppe'] = self.data['AuswKatTyp']
        self.data['AuswGruppe'].replace(self.dic_auswkGruppe, inplace=True)

        self.unique_group = np.sort(self.data['AuswGruppe'].unique())


    def set_flaeche(self, data):

        self.tab_fl_ges = pd.pivot_table(data, values='Fläche in HA', index='Forstrevier', columns=None, aggfunc='sum', fill_value=0, margins=True,)


    ###=========================================================================
    ###   10.x Naturschutz
    ###=========================================================================

    def fuc_tbl_naturschutz(self, typ, merk=0):

        # check if merkmal or typ to filter
        if merk == 0:
            data_filter = self.data[self.data['AuswKatTyp']==typ]
        else:
            data_filter = self.data[self.data['Merkmalausprägung']==merk]

        # get index
        index = self.tab_fl_ges.index

        # pivot table
        tab0 = pd.pivot_table(data_filter, values='Fläche in HA', index='Forstrevier', columns=None, aggfunc='sum', fill_value=0, margins=True,)
        # set index
        tab0 = tab0.reindex(index)
        tab0 = tab0.fillna(0)
        tab0 = tab0.round(1)
        # calc the percentage
        tab1 = (tab0/self.tab_fl_ges)*100
        tab1 = tab1.round(0)
        tab1 = tab1.astype(int)
        # concatonate tables
        table = pd.concat([tab0, tab1, self.tab_fl_ges.round(1)], axis=1)

        # rename columns
        table.columns = ['[ha]', '[%]', '[ha]']
        # get index inside data (for printing into docx)
        # rename index
        table.rename(index={'All':'Ges.'}, inplace=True)

        table.reset_index(level=0, inplace=True)

        return(table)

    def fuc_tbl_biotop(self):

        ## prepare data
        table = self.data[self.data['AuswKatTyp']=='BT']
        # make pivot table
        table = pd.pivot_table(table, values='Fläche in HA', index='Merkmalausprägung', columns='Forstrevier', aggfunc='sum', fill_value=0, margins=True,)

        # get all values from missing data
        uniq =  np.sort(self.data['Forstrevier'].unique()).tolist() + ['All']
        if not table.T.index.values.tolist() == uniq:
            table = self.missing_categories(table.T, uniq).T

        # calc percentage
        table_pre = (table.div(table.iloc[-1,-1], axis=1))*100

        table = table.round(1)
        table_pre_con = round(table_pre,0).astype(int)

        # rename columns
        header = []
        for i in range(table.shape[1]):
            temp = ['[ha]']
            header = header + temp
        table.columns = header
        header = []
        for i in range(table_pre_con.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_pre_con.columns = header

        # make ha and percent fit in one table
        table_all = self.fuc_concat_tables(table, table_pre_con)

        table_all.index.names = ['Biotoptyp']
        table_all.rename(index=self.dic_merk, inplace=True)
        table_all.rename(index={'All':'Gesamtergebnis'}, inplace=True)
        # rename index
        table_all.rename(index={'All':'Ges.'}, inplace=True)

        table_all.reset_index(level=0, inplace=True)

        return table_all


    def fuc_tbl_birdlife(self):

        table = self.data[self.data['AuswKatTyp']=='AHI'][['Forstrevier', 'Abteilung', 'Unterabteil.', 'Teilfl.', 'Bewirtschaftungsform', 'Fläche in HA']]
        table['Unterabteil.'] = table['Unterabteil.'] + table['Teilfl.'].astype(str)
        table['Bewirtschaftungsform'] = table['Bewirtschaftungsform'] + 'W'
        table = table[['Forstrevier', 'Abteilung', 'Unterabteil.', 'Bewirtschaftungsform', 'Fläche in HA']]
        table = table.sort_values(by=['Forstrevier', 'Abteilung', 'Unterabteil.'])
        table.columns = ['Forstrevier', 'Abteilung', 'Waldort', 'Bewirtschaftungsform', 'Fläche [ha]']

        return(table)

    def fuc_tbl_saatgut(self, type):
        return(table)


    ## for concatonating two tables (ha|%)

    def fuc_concat_tables(self,table_base,table_pre):

        # create table_all with first two columns
        table_all = pd.concat([table_base.iloc[:,0], table_pre.iloc[:,0]], axis=1)

        # concat remaining columns to table_all
        for i in np.arange(1,table_base.shape[1]):
            table_all = pd.concat([table_all.iloc[:,:], table_base.iloc[:,i], table_pre.iloc[:,i]], axis=1)

        return(table_all)


    # for missing categories
    def missing_categories(self, table, cat):

        #midx = pd.MultiIndex.from_product([cat], names=['Ertragssituation'])
        table = table.reindex(cat, fill_value=0)

        return table
