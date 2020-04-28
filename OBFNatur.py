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
        self.data_merkmal = data[(data['Merkmalausprägung'] >= 51769) & (data['Merkmalausprägung'] <= 51771)].copy()
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
    ### 9 Schutzwalderhaltungszustand
    ###=========================================================================

    def fuc_tbl_plt_sez_ha(self):

        # create pivot table
        table = pd.pivot_table(self.data_merkmal, index=['Forstrevier'],columns = ['Merkmalausprägung'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        table.columns = table.columns.droplevel()

        # rename
        table.rename(index={'All': 'Ges.'}, inplace=True)
        table.rename(columns={51769:'Grün', 51770:'Gelb', 51771:'Rot', 'All':'Ges.'}, inplace=True)

        # deal with missing columns
        df = pd.DataFrame(columns=['Grün', 'Gelb', 'Rot'])
        table = pd.concat([df,table]).fillna(0)

        ## plot

        # filter sum from column and index
        table_pre = table.loc[table.index != 'Ges.', ['Grün', 'Gelb', 'Rot']]
                # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_pre.plot(kind='bar', stacked=True, ax=ax, width = 0.5, alpha=0.9, title='SW Erhaltungszustand', color=['#108112', '#FFCC2B', '#FC0D1B']) # '#FC0D1B', '#FDB32B', '#FFFD38', '#80DA29', '#108112'
        #edgecolor='white', (possible)
        ax.set_ylabel("Fläche [ha]")

        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=0)

        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        # round
        table = table.round(2)
        table.reset_index(level=0, inplace=True)

        return table


    def fuc_tbl_sez_mass_ha(self, mass, fr):

        # filter Massnahmengruppe
        data_filtered = self.data_merkmal[self.data_merkmal['Massnahmengruppe'] == mass]

        # filter FR
        data_filtered = data_filtered[data_filtered['Forstrevier'] == fr]

        # create pivot table
        table = pd.pivot_table(data_filtered, index=['Maßnahmenart'],columns = ['Merkmalausprägung'], \
                                   values=['Angriffsfläche'], \
                                  aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        table.columns = table.columns.droplevel()

        # rename
        table.rename(index={'All': 'Ges.'}, inplace=True)
        table.rename(columns={51769:'Grün', 51770:'Gelb', 51771:'Rot', 'All':'Ges.'}, inplace=True)

        # deal with missing columns
        df = pd.DataFrame(columns=['Grün', 'Gelb', 'Rot'])
        table = pd.concat([df,table]).fillna(0)

        # round
        table = table.round(1)
        table.reset_index(level=0, inplace=True)

        return table


    def fuc_tbl_sez_mass_v(self, mass, fr):

        # filter Massnahmengruppe
        data_filtered = self.data_merkmal[self.data_merkmal['Massnahmengruppe'] == mass]
        print(data_filtered.shape)

        # filter FR
        data_filtered = data_filtered[data_filtered['Forstrevier'] == fr]
        print(data_filtered.shape)

        # create pivot table
        table_v_lh = pd.pivot_table(data_filtered, index=['Rückungsart'],columns = ['Merkmalausprägung'], \
                                   values=['Nutzung LH'], \
                                  aggfunc=np.sum, fill_value=0, margins=True)

        # create pivot table
        table_v_nh = pd.pivot_table(data_filtered, index=['Rückungsart'],columns = ['Merkmalausprägung'], \
                                   values=['Nutzung NH'], \
                                  aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_v_lh.columns = table_v_lh.columns.droplevel()
        table_v_nh.columns = table_v_nh.columns.droplevel()

        # rename
        table_v_lh.rename(index={'All': 'Ges.'}, inplace=True)
        table_v_nh.rename(index={'All': 'Ges.'}, inplace=True)
        table_v_lh.rename(columns={51769:'Grün', 51770:'Gelb', 51771:'Rot', 'All':'Ges.'}, inplace=True)
        table_v_nh.rename(columns={51769:'Grün', 51770:'Gelb', 51771:'Rot', 'All':'Ges.'}, inplace=True)

        # deal with missing columns
        df = pd.DataFrame(columns=['Grün', 'Gelb', 'Rot'])
        table_v_lh = pd.concat([df,table_v_lh]).fillna(0)
        table_v_nh = pd.concat([df,table_v_nh]).fillna(0)

        # concatonate dataframes
        table_v = pd.concat([table_v_lh.loc[:,'Grün'], table_v_nh.loc[:,'Grün'], \
                table_v_lh.loc[:,'Grün'] + table_v_nh.loc[:,'Grün'], \
                table_v_lh.loc[:,'Gelb'], table_v_nh.loc[:,'Gelb'], \
                table_v_lh.loc[:,'Gelb'] + table_v_nh.loc[:,'Gelb'], \
                table_v_lh.loc[:,'Rot'], table_v_nh.loc[:,'Rot'], \
                table_v_lh.loc[:,'Rot'] + table_v_nh.loc[:,'Rot'], \
                table_v_lh.loc[:,'Ges.'], table_v_nh.loc[:,'Ges.'], \
                table_v_lh.loc[:,'Ges.'] + table_v_nh.loc[:,'Ges.']], axis=1, sort=False)

        table_v = table_v.round(0).astype(int)
        table_v.columns = ['LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']
        table_v.reset_index(level=0, inplace=True)

        return table_v


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
